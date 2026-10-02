// GET /api/admin/data: the full schedule (future weeks too) and every vote with names
import { NAMES, nowWeek, json } from "../../../lib/shared.js";

export async function onRequestGet({ env }) {
  const [sched, votes] = await env.DB.batch([
    env.DB.prepare("SELECT id, year, week, drink, ingredients FROM schedule ORDER BY id"),
    env.DB.prepare("SELECT week_id, name, score, updated_at FROM votes ORDER BY week_id, score DESC, name"),
  ]);
  return json({ now: nowWeek(), names: NAMES, schedule: sched.results, votes: votes.results });
}
