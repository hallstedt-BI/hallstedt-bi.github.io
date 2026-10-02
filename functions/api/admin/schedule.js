// POST /api/admin/schedule { rows: [{year, week, drink, ingredients}], replaceId? }: add or overwrite weeks
// DELETE /api/admin/schedule?id=2026-W40: remove a week (its votes stay but no longer count)
import { wid, weeksIn, json, readJson, clean } from "../../../lib/shared.js";

export async function onRequestPost({ request, env }) {
  const b = await readJson(request);
  const rows = Array.isArray(b?.rows) ? b.rows : [];
  if (!rows.length || rows.length > 200) return json({ error: "Inga veckor att spara." }, 400);
  const stmts = [];
  const t = Date.now();
  for (const r of rows) {
    const y = Number(r.year), w = Number(r.week);
    const drink = clean(r.drink, 80), ingredients = clean(r.ingredients, 400);
    if (!Number.isInteger(y) || y < 2000 || y > 2100 || !Number.isInteger(w) || w < 1 || w > weeksIn(y) || !drink)
      return json({ error: `Ogiltig rad: vecka ${r.week} ${r.year}.` }, 400);
    stmts.push(env.DB.prepare(
      `INSERT INTO schedule (id, year, week, drink, ingredients, updated_at) VALUES (?, ?, ?, ?, ?, ?)
       ON CONFLICT (id) DO UPDATE SET drink = excluded.drink, ingredients = excluded.ingredients, updated_at = excluded.updated_at`
    ).bind(wid(y, w), y, w, drink, ingredients, t));
  }
  // Moving an edited week to a new week number: drop the old row and carry its votes along
  if (b.replaceId && rows.length === 1) {
    const newId = wid(Number(rows[0].year), Number(rows[0].week));
    if (b.replaceId !== newId) {
      stmts.push(env.DB.prepare("UPDATE OR REPLACE votes SET week_id = ? WHERE week_id = ?").bind(newId, b.replaceId));
      stmts.push(env.DB.prepare("DELETE FROM schedule WHERE id = ?").bind(b.replaceId));
    }
  }
  await env.DB.batch(stmts);
  return json({ ok: true, saved: rows.length });
}

export async function onRequestDelete({ request, env }) {
  const id = new URL(request.url).searchParams.get("id");
  if (!id) return json({ error: "id saknas" }, 400);
  await env.DB.prepare("DELETE FROM schedule WHERE id = ?").bind(id).run();
  return json({ ok: true });
}
