// ─────────────────────────────────────────────────────────────
//  common.js — shared helpers for the tournament system
//  Loaded by index.html, admin.html and betting.html.
//  Pure functions only (no Firebase calls) so they're easy to reason about.
// ─────────────────────────────────────────────────────────────

const TOTAL_TOKENS = 10;     // tokens each bettor starts with
const MIN_ODDS = 1.2;        // floor so even safe bets pay something
const MAX_ODDS = 500;        // cap (also enforced by DB rules)

// ── Team visual identity: deterministic color + initials ──
const TEAM_PALETTE = [
    '#ff6b6b', '#2ec4b6', '#4d96ff', '#ffa94d', '#9775fa',
    '#f783ac', '#20c997', '#fcc419', '#5c7cfa', '#ff8787',
    '#38d9a9', '#e599f7', '#ff922b', '#74c0fc', '#b197fc',
    '#63e6be'
];

function teamColor(name) {
    if (!name) return '#aaa';
    let hash = 0;
    for (let i = 0; i < name.length; i++) {
        hash = (hash * 31 + name.charCodeAt(i)) % 1000000007;
    }
    return TEAM_PALETTE[Math.abs(hash) % TEAM_PALETTE.length];
}

function teamInitials(name) {
    if (!name) return '?';
    // Split on / or whitespace, take first letter of up to 2 parts
    const parts = name.split(/[\/\s]+/).filter(Boolean);
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
    return (parts[0][0] + parts[1][0]).toUpperCase();
}

// Returns an HTML string for a colored badge with initials
function teamBadge(name, size) {
    if (!name) return '';
    const px = size || 22;
    const fs = Math.round(px * 0.42);
    const color = teamColor(name);
    return `<span class="team-badge" style="width:${px}px;height:${px}px;background:${color};font-size:${fs}px;">${teamInitials(name)}</span>`;
}

// ── Group stage standings ──
// matches: array of { home, away, homeScore, awayScore } (scores null/'' if unplayed)
// Returns array sorted by points, GD, GF, name
function calculateStandings(matches, teamNames) {
    const teams = {};
    function ensure(name) {
        if (!teams[name]) teams[name] = { name, played: 0, wins: 0, draws: 0, losses: 0, gf: 0, ga: 0, gd: 0, points: 0 };
    }
    (teamNames || []).forEach(ensure);

    matches.forEach(m => {
        ensure(m.home); ensure(m.away);
        const hs = m.homeScore, as = m.awayScore;
        if (hs == null || hs === '' || as == null || as === '') return;
        const h = parseInt(hs), a = parseInt(as);
        if (isNaN(h) || isNaN(a)) return;

        const home = teams[m.home], away = teams[m.away];
        home.played++; away.played++;
        home.gf += h; home.ga += a;
        away.gf += a; away.ga += h;

        if (h > a) { home.wins++; away.losses++; home.points += 3; }
        else if (h < a) { away.wins++; home.losses++; away.points += 3; }
        else { home.draws++; away.draws++; home.points++; away.points++; }
    });

    Object.values(teams).forEach(t => t.gd = t.gf - t.ga);
    return Object.values(teams).sort((a, b) =>
        b.points - a.points || b.gd - a.gd || b.gf - a.gf || a.name.localeCompare(b.name)
    );
}

// ── Team attack/defense strengths (for odds) ──
// Uses Laplace smoothing (+1 virtual goal) so perfect-defense / no-attack
// teams don't collapse the model to 0% and produce Infinity odds.
function calculateTeamStats(matches) {
    const teams = {};
    let totalGoals = 0, totalGames = 0;
    function ensure(name) {
        if (!teams[name]) teams[name] = { name, played: 0, gf: 0, ga: 0 };
    }

    matches.forEach(m => {
        ensure(m.home); ensure(m.away);
        const hs = m.homeScore, as = m.awayScore;
        if (hs == null || hs === '' || as == null || as === '') return;
        const h = parseInt(hs), a = parseInt(as);
        if (isNaN(h) || isNaN(a)) return;
        totalGoals += h + a; totalGames++;
        teams[m.home].played++; teams[m.away].played++;
        teams[m.home].gf += h; teams[m.home].ga += a;
        teams[m.away].gf += a; teams[m.away].ga += h;
    });

    const leagueAvg = totalGames > 0 ? totalGoals / (totalGames * 2) : 1;
    Object.values(teams).forEach(t => {
        if (t.played > 0) {
            t.attack = ((t.gf + 1) / (t.played + 1)) / leagueAvg;
            t.defense = ((t.ga + 1) / (t.played + 1)) / leagueAvg;
        } else {
            t.attack = 1; t.defense = 1;
        }
    });

    return { teams, leagueAvg };
}

// ── Poisson helpers ──
function logFactorial(n) {
    let s = 0;
    for (let i = 2; i <= n; i++) s += Math.log(i);
    return s;
}
function poissonPmf(k, lambda) {
    if (lambda <= 0) return k === 0 ? 1 : 0;
    return Math.exp(-lambda + k * Math.log(lambda) - logFactorial(k));
}

// Win odds for a knockout match (no draws — draw mass redistributed)
// Returns { homeWin, awayWin, homeProb, awayProb }
function getMatchOdds(stats, homeTeam, awayTeam) {
    const home = stats.teams[homeTeam], away = stats.teams[awayTeam];
    if (!home || !away || !stats.leagueAvg) return { homeWin: 2, awayWin: 2, homeProb: 0.5, awayProb: 0.5 };

    const avg = stats.leagueAvg;
    const expHome = home.attack * away.defense * avg;
    const expAway = away.attack * home.defense * avg;

    let homeWin = 0, awayWin = 0;
    for (let i = 0; i <= 10; i++) {
        for (let j = 0; j <= 10; j++) {
            const p = poissonPmf(i, expHome) * poissonPmf(j, expAway);
            if (i > j) homeWin += p;
            else if (i < j) awayWin += p;
        }
    }
    const total = homeWin + awayWin;
    if (total === 0) return { homeWin: 2, awayWin: 2, homeProb: 0.5, awayProb: 0.5 };
    homeWin /= total; awayWin /= total;

    return {
        homeWin: clampOdds(1 / Math.max(homeWin, 1 / MAX_ODDS)),
        awayWin: clampOdds(1 / Math.max(awayWin, 1 / MAX_ODDS)),
        homeProb: homeWin,
        awayProb: awayWin
    };
}

// First-goal-scorer odds for every player on both teams.
// rosters: { teamName: [player, ...] }
// playerGoals: { player: totalGoals }
// Laplace smoothing so non-scorers still appear (with long odds).
function getFirstGoalOdds(stats, rosters, playerGoals, homeTeam, awayTeam) {
    const odds = getMatchOdds(stats, homeTeam, awayTeam);
    const out = [];

    function addTeam(team, teamProb) {
        const players = rosters[team] || [];
        const total = players.reduce((s, p) => s + (playerGoals[p] || 0) + 1, 0) || 1;
        players.forEach(p => {
            const share = ((playerGoals[p] || 0) + 1) / total;
            const prob = teamProb * share;
            out.push({ name: p, team, prob, odds: clampOdds(1 / Math.max(prob, 1 / MAX_ODDS)) });
        });
    }
    addTeam(homeTeam, odds.homeProb);
    addTeam(awayTeam, odds.awayProb);
    return out.sort((a, b) => a.odds - b.odds);
}

function clampOdds(raw) {
    return Math.min(MAX_ODDS, Math.max(MIN_ODDS, Math.round(raw * 10) / 10));
}

// ── Round-robin schedule (every team plays every other once) ──
// teamNames: array. Returns array of { home, away, homeScore:null, awayScore:null, playedOrder:null, goals:{} }
function generateRoundRobin(teamNames) {
    const matches = [];
    for (let i = 0; i < teamNames.length; i++) {
        for (let j = i + 1; j < teamNames.length; j++) {
            matches.push({
                home: teamNames[i],
                away: teamNames[j],
                homeScore: null,
                awayScore: null,
                playedOrder: null,
                goals: {}
            });
        }
    }
    return matches;
}

// ── Playoff bracket from final standings ──
// standings: sorted array (from calculateStandings). qualifiers: 4 or 8.
// Returns an object keyed by match id, each match carries advancesTo/advancesSlot
// pointers so winners auto-feed the next round.
function generateBracket(standings, qualifiers) {
    const seeds = standings.slice(0, qualifiers).map(t => t.name);
    const matches = {};

    if (qualifiers === 4) {
        // SF1: 1v4, SF2: 2v3, Final
        matches.sf1 = mkMatch('sf1', 'Semifinal 1', 'sf', seeds[0], seeds[3], 'final', 'home');
        matches.sf2 = mkMatch('sf2', 'Semifinal 2', 'sf', seeds[1], seeds[2], 'final', 'away');
        matches.final = mkMatch('final', 'Final', 'final', null, null, null, null);
    } else if (qualifiers === 8) {
        // Standard 8-seed bracket so seeds 1 & 2 can only meet in the final
        matches.qf1 = mkMatch('qf1', 'Quarterfinal 1', 'qf', seeds[0], seeds[7], 'sf1', 'home');
        matches.qf2 = mkMatch('qf2', 'Quarterfinal 2', 'qf', seeds[3], seeds[4], 'sf1', 'away');
        matches.qf3 = mkMatch('qf3', 'Quarterfinal 3', 'qf', seeds[1], seeds[6], 'sf2', 'home');
        matches.qf4 = mkMatch('qf4', 'Quarterfinal 4', 'qf', seeds[2], seeds[5], 'sf2', 'away');
        matches.sf1 = mkMatch('sf1', 'Semifinal 1', 'sf', null, null, 'final', 'home');
        matches.sf2 = mkMatch('sf2', 'Semifinal 2', 'sf', null, null, 'final', 'away');
        matches.final = mkMatch('final', 'Final', 'final', null, null, null, null);
    }
    return matches;
}

function mkMatch(id, label, round, home, away, advancesTo, advancesSlot) {
    return {
        id, label, round,
        home: home || null,
        away: away || null,
        homeScore: null, awayScore: null,
        winner: null, firstGoalScorer: null,
        locked: false, completed: false,
        goals: {},
        advancesTo: advancesTo || null,
        advancesSlot: advancesSlot || null
    };
}

// Ordered list of match ids for display (left→right rounds)
function bracketOrder(qualifiers) {
    if (qualifiers === 4) return ['sf1', 'sf2', 'final'];
    if (qualifiers === 8) return ['qf1', 'qf2', 'qf3', 'qf4', 'sf1', 'sf2', 'final'];
    return [];
}

// Resolve the winner of a completed match (respects manual penalty winner)
function matchWinner(m) {
    if (!m || !m.completed) return null;
    if (m.winner) return m.winner;
    if (m.homeScore == null || m.awayScore == null) return null;
    if (m.homeScore === m.awayScore) return null; // needs manual winner
    return m.homeScore > m.awayScore ? m.home : m.away;
}

// Sum every player's goals across a set of match objects (each match.goals = {player:count})
function tallyPlayerGoals(matchArrays) {
    const totals = {};
    matchArrays.forEach(arr => {
        (arr || []).forEach(m => {
            if (!m || !m.goals) return;
            Object.entries(m.goals).forEach(([player, n]) => {
                totals[player] = (totals[player] || 0) + (parseInt(n) || 0);
            });
        });
    });
    return totals;
}

// Normalize a Firebase node that may be an array or an object-keyed map into an array
function toArray(data) {
    if (!data) return [];
    if (Array.isArray(data)) return data.filter(Boolean);
    return Object.keys(data).map(k => data[k]);
}
