// GET /api/state: public data. Schedule up to this week, vote counts and sums. No names of voters.
import { NAMES, OPEN_H, OPEN_M, nowWeek, opensAt, json } from "../../lib/shared.js";

export async function onRequestGet({ env }) {
  const now = nowWeek();
  const [sched, agg] = await env.DB.batch([
    env.DB.prepare("SELECT id, year, week, drink, ingredients FROM schedule WHERE id <= ? ORDER BY id").bind(now.id),
    env.DB.prepare("SELECT week_id, COUNT(*) AS n, SUM(score) AS s FROM votes GROUP BY week_id"),
  ]);
  const byWeek = Object.fromEntries(agg.results.map((r) => [r.week_id, r]));
  return json({
    now, serverTime: Date.now(), opens: { hour: OPEN_H, minute: OPEN_M }, names: NAMES,
    weeks: sched.results.map((e) => ({
      ...e, opensAt: opensAt(e.year, e.week),
      votes: byWeek[e.id]?.n || 0, sum: byWeek[e.id]?.s || 0,
    })),
  });
}
