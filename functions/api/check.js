// POST /api/check { passcode }: lets the phone confirm the office password before showing the vote steps
import { json, readJson, safeEqual } from "../../lib/shared.js";

export async function onRequestPost({ request, env }) {
  const body = await readJson(request);
  if (!env.OFFICE_PASSWORD) return json({ error: "OFFICE_PASSWORD is not configured" }, 500);
  if (!body || !safeEqual(body.passcode, env.OFFICE_PASSWORD)) return json({ ok: false }, 401);
  return json({ ok: true });
}
