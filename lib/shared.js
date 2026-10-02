// Shared server-side helpers for Fredagsdrinken (Cloudflare Pages Functions)

// The people who can vote. Edit this list and redeploy to add or remove someone.
export const NAMES = [
  "Amanda", "Martin F", "Martin H", "Jonas", "George", "Mark", "Stephanie",
  "Julia", "Anders K", "Anders G", "Mood", "Kristian", "Silvan", "Daniel",
  "Emil", "Simon", "Louise", "Dara", "Andrea", "Aleks",
];

// Voting for a week opens on that week's Friday at this Stockholm time
export const OPEN_H = 15;
export const OPEN_M = 40;

const pad = (n) => String(n).padStart(2, "0");
export const wid = (y, w) => `${y}-W${pad(w)}`;

export function isoWeekFromYMD(y, m, d) {
  const t = new Date(Date.UTC(y, m, d));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const yy = t.getUTCFullYear();
  return { year: yy, week: Math.ceil(((t - Date.UTC(yy, 0, 1)) / 86400000 + 1) / 7) };
}
export const weeksIn = (y) => isoWeekFromYMD(y, 11, 28).week;

export function fridayOf(y, w) {
  const jan4 = new Date(Date.UTC(y, 0, 4));
  const d = new Date(jan4);
  d.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay() || 7) - 1) + (w - 1) * 7 + 4);
  return d;
}

const sthlm = new Intl.DateTimeFormat("en-US", {
  timeZone: "Europe/Stockholm", hourCycle: "h23",
  year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit",
});
const partsOf = (t) => Object.fromEntries(sthlm.formatToParts(new Date(t)).map((x) => [x.type, x.value]));

export function stockholmToMs(y, mo, d, h, mi) {
  const want = Date.UTC(y, mo, d, h, mi);
  let t = want;
  for (let i = 0; i < 3; i++) {
    const p = partsOf(t);
    t += want - Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute);
  }
  return t;
}
export function opensAt(y, w) {
  const f = fridayOf(y, w);
  return stockholmToMs(f.getUTCFullYear(), f.getUTCMonth(), f.getUTCDate(), OPEN_H, OPEN_M);
}
export function nowWeek(now = Date.now()) {
  const p = partsOf(now);
  const w = isoWeekFromYMD(+p.year, +p.month - 1, +p.day);
  return { ...w, id: wid(w.year, w.week) };
}

export const json = (data, status = 200) =>
  new Response(JSON.stringify(data), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" },
  });

export function safeEqual(a, b) {
  a = String(a ?? ""); b = String(b ?? "");
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

export async function readJson(request) {
  try { return await request.json(); } catch { return null; }
}

export const clean = (s, max) => String(s ?? "").replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
