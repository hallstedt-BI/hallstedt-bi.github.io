// POST /api/vote { passcode, weekId, name, score }: one vote per name per week; voting again replaces it
import { NAMES, nowWeek, opensAt, json, readJson, safeEqual } from "../../lib/shared.js";

export async function onRequestPost({ request, env }) {
  const b = await readJson(request);
  if (!b) return json({ error: "Ogiltig förfrågan." }, 400);
  if (!env.OFFICE_PASSWORD || !safeEqual(b.passcode, env.OFFICE_PASSWORD)) return json({ error: "Fel lösenord.", code: "bad_pass" }, 401);
  if (!NAMES.includes(b.name)) return json({ error: "Välj ditt namn i listan." }, 400);
  const score = Number(b.score);
  if (!Number.isInteger(score) || score < 1 || score > 10) return json({ error: "Betyget måste vara 1 till 10." }, 400);

  const week = await env.DB.prepare("SELECT id, year, week FROM schedule WHERE id = ?").bind(String(b.weekId || "")).first();
  if (!week || week.id > nowWeek().id) return json({ error: "Den veckan finns inte att rösta på." }, 400);
  if (Date.now() < opensAt(week.year, week.week)) return json({ error: "Röstningen har inte öppnat än. Den öppnar fredag kl. 15:40.", code: "not_open" }, 403);

  const t = Date.now();
  await env.DB.prepare(
    `INSERT INTO votes (week_id, name, score, created_at, updated_at) VALUES (?, ?, ?, ?, ?)
     ON CONFLICT (week_id, name) DO UPDATE SET score = excluded.score, updated_at = excluded.updated_at`
  ).bind(week.id, b.name, score, t, t).run();
  return json({ ok: true });
}
