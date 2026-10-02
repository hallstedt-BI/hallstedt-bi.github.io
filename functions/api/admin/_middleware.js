// Every /api/admin/* request needs the committee password in the x-admin-key header
import { json, safeEqual } from "../../../lib/shared.js";

export async function onRequest({ request, env, next }) {
  if (!env.ADMIN_PASSWORD) return json({ error: "ADMIN_PASSWORD is not configured" }, 500);
  if (!safeEqual(request.headers.get("x-admin-key"), env.ADMIN_PASSWORD)) return json({ error: "Fel kommittélösenord." }, 401);
  return next();
}
