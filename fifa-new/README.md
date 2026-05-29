# 🏆 Tournament + Betting System

A self-contained, no-backend tournament tracker with a built-in betting game. Runs as
plain static HTML — host it anywhere (GitHub Pages, Netlify, a USB stick). All live data
lives in a free Firebase Realtime Database.

What you get:

- **Group stage** — round-robin table, auto-generated schedule, live standings
- **Playoff bracket** — single-elimination, seeded from the group standings (4 or 8 teams)
- **Top-scorer list** — goals entered inline with each match result
- **Betting game** — everyone gets 10 tokens to bet on match winners and first-goal scorers,
  with odds calculated automatically from group-stage performance
- **Admin panel** — set up teams, generate the schedule + bracket, enter results, manage bettors

Everything is white-themed and works on a TV/projector for live viewing.

---

## Files

| File | What it is |
|------|-----------|
| `index.html` | Public view — group table, bracket, top scorers, token leaderboard |
| `betting.html` | Where people sign up and place bets |
| `admin.html` | Admin-only — setup, results entry, bettor management |
| `common.js` | Shared logic (standings, odds, schedule/bracket generation) |
| `firebase-config.js` | **You edit this** — your Firebase project keys |
| `README.md` | This file |

---

## Setup (about 10 minutes)

### 1. Create a Firebase project

1. Go to the [Firebase Console](https://console.firebase.google.com) → **Add project**. Give it any name.
2. In the left sidebar: **Build → Realtime Database → Create Database**.
   - Pick a location (e.g. `europe-west1`).
   - Start in **test mode** (we lock it down in step 3).
3. Click the gear icon → **Project settings → General**. Scroll to **Your apps** and click
   the web icon (`</>`). Register an app (any nickname). Firebase shows you a `firebaseConfig`
   object — copy those values.

### 2. Paste your config

Open `firebase-config.js` and replace the placeholders with the values from step 1:

```js
const firebaseConfig = {
    apiKey: "…",
    authDomain: "your-project.firebaseapp.com",
    databaseURL: "https://your-project-default-rtdb.europe-west1.firebasedatabase.app",
    projectId: "your-project",
    storageBucket: "your-project.appspot.com",
    messagingSenderId: "…",
    appId: "…"
};
```

> The `apiKey` here is **safe to be public** — it's a client identifier, not a secret. Security
> is enforced by the database rules in step 3. (GitHub may warn about a "Google API key" in this
> file; you can dismiss it.)

### 3. Set the database rules

In the Firebase Console: **Realtime Database → Rules**. Replace everything with the block below
and click **Publish**. These rules make all data publicly *readable* (so the bracket can be shown
without logging in), but tightly control *writes*:

```json
{
  "rules": {
    "config": {
      ".read": true,
      ".write": "auth != null && (!data.child('adminEmail').exists() || data.child('adminEmail').val() == auth.token.email)"
    },
    "teams": {
      ".read": true,
      ".write": "auth != null && auth.token.email == root.child('config/adminEmail').val()"
    },
    "groupStage": {
      ".read": true,
      ".write": "auth != null && auth.token.email == root.child('config/adminEmail').val()"
    },
    "playoff": {
      ".read": true,
      ".write": "auth != null && auth.token.email == root.child('config/adminEmail').val()"
    },
    "users": {
      ".read": true,
      "$name": {
        ".write": "auth != null && (!data.exists() || data.child('uid').val() == auth.uid || auth.token.email == root.child('config/adminEmail').val())"
      }
    },
    "bets": {
      ".read": true,
      "$matchId": {
        "$name": {
          ".write": "auth != null && (root.child('users').child($name).child('uid').val() == auth.uid || auth.token.email == root.child('config/adminEmail').val()) && root.child('playoff/matches').child($matchId).child('locked').val() != true && root.child('playoff/matches').child($matchId).child('completed').val() != true",
          "$betKey": {
            ".validate": "newData.hasChildren(['type','pick','amount','odds']) && newData.child('amount').isNumber() && newData.child('amount').val() > 0 && newData.child('amount').val() <= 1000 && newData.child('odds').isNumber() && newData.child('odds').val() >= 1 && newData.child('odds').val() <= 500"
        }
        }
      }
    }
  }
}
```

What these rules guarantee:

- Anyone can **view** the tournament (read-only).
- Only the **admin** can edit teams, schedule, results and the bracket.
- A logged-in user can only **edit their own** account and bets.
- Bets can only be placed on matches that are **open** (not locked or finished) — no betting
  after kickoff, no fabricated odds.

### 4. Turn on email/password login

Firebase Console: **Build → Authentication → Get started → Email/Password → Enable → Save**.

### 5. Claim the admin account

1. Open `betting.html` in a browser, click **Sign up**, and register with your email + a password
   and a username. (This becomes your personal account.)
2. Open `admin.html`, log in with that same email/password.
3. Because no admin exists yet, you'll see a **Claim admin** banner — click it. From now on, only
   that email can administer the tournament.

> Want a different admin email than your bettor account? Just sign up with that email on the betting
> page first, then log into admin with it and claim.

### 6. Deploy

It's all static files. Options:

- **GitHub Pages** — push this folder to a repo, enable Pages, done.
- **Netlify / Vercel** — drag-and-drop the folder.
- **Local** — run `python3 -m http.server` in this folder and open `http://localhost:8000`.
  (Don't open the files with `file://` — Firebase needs `http(s)://`.)

---

## Running a tournament

### Before it starts (admin → Setup tab)

1. **Tournament** — set a name and choose how many teams advance to the playoff (**4** or **8**).
2. **Teams** — add each team and its players (comma-separated). Players power the top-scorer
   list and first-goal bets. Solo tournament? Just put one player per team.
3. **Generate group schedule** — creates the round-robin fixtures.

### During the group stage (admin → Group stage tab)

- For each match, enter the score and (optionally) how many goals each player scored.
  Goals feed the top-scorer list. Click **Save**.
- The public `index.html` updates live.

### Starting the playoff (admin → Setup tab)

- Once the group is done, click **Generate bracket from standings**. The top 4 (or 8) teams are
  seeded automatically (1 vs lowest, etc.).
- People can now bet on the first round at `betting.html`.

### Each playoff match (admin → Playoff tab)

1. Click **Lock bets & start** when the match is about to begin — this closes betting.
2. Enter the score + goals + the first-goal scorer.
   - **Draw / penalties?** Enter the regulation score (e.g. 1–1) and pick the **winner** in the
     dropdown. The winner advances on penalties.
3. Click **Save result** — the winner moves into the next round automatically.
4. Made a mistake? **Reopen** re-opens the match for editing.

### Betting (everyone, on `betting.html`)

- Each person signs up once (email + password + username) and gets **10 tokens**.
- Click any open match to bet on the **winner** or **first goal scorer**. Odds come from group-stage
  form (a Poisson model). You can split tokens across matches or go all-in.
- Win → your payout (stake × odds) is added to your balance and can be re-bet on later rounds.
  Lose → the staked tokens are gone.
- Bets can be changed or removed until the match is locked.
- Odds for later rounds update after each result is entered.

---

## How the odds work

Each team gets an **attack** and **defense** rating from group-stage goals (scored / conceded per
game, relative to the tournament average, with +1 smoothing so a perfect defense doesn't break the
math). A Poisson model turns those into win probabilities for any matchup; draw probability is
redistributed since knockouts can't end level. First-goal odds combine the team's chance of scoring
with each player's share of their team's goals. Odds have a floor of **1.2x** (even safe bets pay
something) and a cap of **500x**.

---

## Customizing

- **Colors** — the theme uses `#ff6b6b` (red) and `#2ec4b6` (teal). Search-and-replace in the
  three HTML files to re-skin.
- **Starting tokens / odds floor & cap** — top of `common.js` (`TOTAL_TOKENS`, `MIN_ODDS`, `MAX_ODDS`).
- **Team colors** — auto-assigned from `TEAM_PALETTE` in `common.js`.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| Login spins / `permission_denied` in console | Rules not published, or Email/Password auth not enabled (steps 3–4). |
| "This account is not the admin" | You're logged into admin with a different email than the one that claimed admin. |
| Bracket shows "TBD" | Earlier-round matches aren't completed yet — winners haven't been decided. |
| Can't place a bet | The match is locked/finished, or you're out of tokens. |
| Odds show a huge number | Expected for heavy underdogs (capped at 500x). |
| Changed teams after generating schedule | Re-generate the schedule (it clears old results). |

---

Built as an open, reusable version of a friends-and-family FIFA tournament tracker. Have fun. 🎮
