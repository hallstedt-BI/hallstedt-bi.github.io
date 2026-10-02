// DELETE /api/admin/vote?week=2026-W40&name=Anna: remove a vote someone shouldn't have made
import { json } from "../../../lib/shared.js";

export async function onRequestDelete({ request, env }) {
  const q = new URL(request.url).searchParams;
  const week = q.get("week"), name = q.get("name");
  if (!week || !name) return json({ error: "week och name krävs" }, 400);
  await env.DB.prepare("DELETE FROM votes WHERE week_id = ? AND name = ?").bind(week, name).run();
  return json({ ok: true });
}
