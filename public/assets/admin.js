// Committee page: log in with the committee password, edit the schedule, import from Excel, see and remove votes
(() => {
  "use strict";
  const { el, $, api, store, wid } = FD;
  const KEY = "fd-admin";
  const NOW = FD.isoWeek(new Date());
  const NOW_ID = wid(NOW.year, NOW.week);
  const S = { key: store.get(KEY), schedule: [], votes: [], names: [], year: NOW.year, open: new Set(), editing: null };

  const adminApi = (path, opts = {}) => api(path, { ...opts, headers: { "x-admin-key": S.key || "" } });

  // ---------- login ----------
  async function tryLogin(key) {
    S.key = key;
    try { await refresh(); store.set(KEY, key); $("login").hidden = true; $("panel").hidden = false; fillForm(null); return true; }
    catch (e) { S.key = null; store.del(KEY); $("login").hidden = false; $("panel").hidden = true; return e; }
  }
  $("admin-go").addEventListener("click", async () => {
    const v = $("admin-pass").value.trim(); if (!v) return;
    const r = await tryLogin(v);
    if (r !== true) $("login-msg").textContent = r.status === 401 ? "Fel lösenord." : "Kunde inte logga in. Försök igen.";
  });
  $("admin-pass").addEventListener("keydown", (e) => { if (e.key === "Enter") $("admin-go").click(); });
  $("logout").addEventListener("click", () => { store.del(KEY); location.reload(); });

  async function refresh() {
    const d = await adminApi("api/admin/data");
    S.schedule = d.schedule; S.votes = d.votes; S.names = d.names;
    render();
  }

  // ---------- single week form ----------
  function fillForm(e) {
    S.editing = e ? e.id : null;
    $("f-year").value = e ? e.year : S.year;
    $("f-week").value = e ? e.week : "";
    $("f-drink").value = e ? e.drink : "";
    $("f-note").value = e?.ingredients || "";
    $("form-title").textContent = e ? `Ändra vecka ${e.week}` : "Lägg till vecka";
    $("f-save").textContent = e ? "Spara ändringen" : "Spara vecka";
    $("f-cancel").hidden = !e;
    $("f-msg").textContent = "";
  }
  async function saveWeek() {
    const y = parseInt($("f-year").value, 10), w = parseInt($("f-week").value, 10);
    const drink = $("f-drink").value.trim(), ingredients = $("f-note").value.trim(), msg = $("f-msg");
    if (!(y >= 2000 && y <= 2100)) { msg.textContent = `Ange ett år, till exempel ${NOW.year}.`; return; }
    if (!(w >= 1 && w <= FD.weeksIn(y))) { msg.textContent = `Vecka måste vara mellan 1 och ${FD.weeksIn(y)} för ${y}.`; return; }
    if (!drink) { msg.textContent = "Skriv in vilken drink som serveras."; $("f-drink").focus(); return; }
    $("f-save").disabled = true;
    try {
      await adminApi("api/admin/schedule", { method: "POST", body: { rows: [{ year: y, week: w, drink, ingredients }], replaceId: S.editing } });
      S.year = y; fillForm(null); msg.textContent = `Sparat: vecka ${w}, ${drink}.`; await refresh(); $("f-week").focus();
    } catch (e) { msg.textContent = e.message || "Det gick inte att spara. Försök igen."; }
    finally { $("f-save").disabled = false; }
  }
  $("f-save").addEventListener("click", saveWeek);
  $("f-cancel").addEventListener("click", () => fillForm(null));
  for (const id of ["f-year", "f-week", "f-drink", "f-note"]) $(id).addEventListener("keydown", (e) => { if (e.key === "Enter") saveWeek(); });

  // ---------- paste from Excel ----------
  function parsePaste(text, defYear) {
    const num = (s) => { const m = String(s).match(/^(?:v(?:ecka|\.)?\s*)?(\d{1,4})$/i); return m ? parseInt(m[1], 10) : null; };
    const out = new Map();
    for (const raw of text.split(/\r?\n/)) {
      const line = raw.trim(); if (!line) continue;
      let cells = raw.split("\t").map((s) => s.trim()).filter(Boolean);
      if (cells.length < 2) cells = line.split(/\s*;\s*/).filter(Boolean);
      if (cells.length < 2) { const m = line.match(/^(?:v(?:ecka|\.)?\s*)?(\d{1,2})\s+(.+)$/i); if (m) cells = [m[1], m[2]]; }
      if (cells.length < 2) continue;
      let y = defYear, w = null, drink = null, ingredients = "";
      const date = cells[0].match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (date) { const iw = FD.isoWeek(new Date(+date[1], +date[2] - 1, +date[3])); y = iw.year; w = iw.week; drink = cells[1]; ingredients = cells.slice(2).join(", "); }
      else if (num(cells[0]) >= 2000 && cells.length >= 3 && num(cells[1])) { y = num(cells[0]); w = num(cells[1]); drink = cells[2]; ingredients = cells.slice(3).join(", "); }
      else if (num(cells[0]) && num(cells[0]) <= 53) { w = num(cells[0]); drink = cells[1]; ingredients = cells.slice(2).join(", "); }
      if (!w || !drink || w > FD.weeksIn(y)) continue;
      out.set(wid(y, w), { year: y, week: w, drink: drink.slice(0, 80), ingredients: ingredients.slice(0, 400) });
    }
    return [...out.values()];
  }
  function updatePreview() {
    const rows = parsePaste($("p-text").value, parseInt($("f-year").value, 10) || S.year);
    const btn = $("p-import"), pv = $("p-preview");
    btn.disabled = !rows.length;
    btn.textContent = rows.length ? `Importera ${rows.length} ${rows.length === 1 ? "vecka" : "veckor"}` : "Importera";
    pv.style.whiteSpace = "pre-line";
    pv.textContent = rows.length
      ? rows.slice(0, 6).map((r) => `v. ${r.week} ${r.year}: ${r.drink}${r.ingredients ? " (" + (r.ingredients.length > 40 ? r.ingredients.slice(0, 40) + "…" : r.ingredients) + ")" : ""}`).join("\n") + (rows.length > 6 ? `\n…och ${rows.length - 6} till` : "")
      : ($("p-text").value.trim() ? "Hittar inga rader med vecka och drink. Kontrollera att första kolumnen är veckonummer." : "");
    return rows;
  }
  $("p-text").addEventListener("input", updatePreview);
  $("f-year").addEventListener("input", updatePreview);
  $("p-import").addEventListener("click", async () => {
    const rows = updatePreview(); if (!rows.length) return;
    const btn = $("p-import"), msg = $("p-msg"); btn.disabled = true; msg.textContent = "Importerar…";
    try {
      for (let i = 0; i < rows.length; i += 100) await adminApi("api/admin/schedule", { method: "POST", body: { rows: rows.slice(i, i + 100) } });
      msg.textContent = `Importerade ${rows.length} ${rows.length === 1 ? "vecka" : "veckor"}.`;
      S.year = rows[0].year; $("p-text").value = ""; updatePreview(); await refresh();
    } catch (e) { msg.textContent = e.message || "Importen misslyckades. Försök igen."; btn.disabled = false; }
  });

  // ---------- schedule list with votes ----------
  $("y-prev").addEventListener("click", () => { S.year--; render(); });
  $("y-next").addEventListener("click", () => { S.year++; render(); });

  function render() {
    $("y-label").textContent = S.year;
    $("tv-qr").innerHTML = FD.qrSvg(FD.voteUrl());
    $("names").textContent = `Röstberättigade (${S.names.length}): ${S.names.join(", ")}. Listan ändras i lib/shared.js.`;
    const ul = $("sched"); ul.replaceChildren();
    const list = S.schedule.filter((e) => e.year === S.year);
    if (!list.length) { ul.append(el("li", { class: "sched-empty", text: `Inga veckor inlagda för ${S.year}. Lägg till en vecka eller klistra in från Excel.` })); return; }
    for (const e of list) {
      const vs = S.votes.filter((v) => v.week_id === e.id);
      const avg = vs.length ? FD.fmtAvg(Math.round((vs.reduce((a, v) => a + v.score, 0) / vs.length) * 10) / 10) : null;
      const open = S.open.has(e.id);
      const del = el("button", { type: "button", class: "link danger", text: "Ta bort", "data-armed": "false" });
      del.addEventListener("click", async () => {
        if (del.dataset.armed !== "true") { del.dataset.armed = "true"; del.textContent = "Bekräfta borttagning"; setTimeout(() => { del.dataset.armed = "false"; del.textContent = "Ta bort"; }, 4000); return; }
        try { await adminApi(`api/admin/schedule?id=${encodeURIComponent(e.id)}`, { method: "DELETE" }); await refresh(); } catch { del.textContent = "Kunde inte ta bort"; }
      });
      const li = el("li", { class: "sch" + (e.id === NOW_ID ? " now" : "") },
        el("div", { class: "sch-main" },
          el("span", { class: "wk", text: `v. ${e.week}` }),
          el("span", { class: "date", text: FD.fmtDate(FD.fridayOf(e.year, e.week)) }),
          el("span", { class: "dn", text: e.drink }),
          el("span", { class: "sch-actions" },
            el("button", { type: "button", class: "link", "aria-expanded": String(open), text: vs.length ? `${vs.length} ${vs.length === 1 ? "röst" : "röster"}, snitt ${avg}` : "Inga röster",
              onclick: () => { open ? S.open.delete(e.id) : S.open.add(e.id); render(); } }),
            el("button", { type: "button", class: "link", text: "Ändra", onclick: () => { fillForm(e); $("f-drink").focus(); } }),
            del)));
      if (open) {
        const box = el("div", { class: "voters" });
        if (!vs.length) box.append(el("span", { class: "none", text: "Ingen har röstat på den här veckan än." }));
        for (const v of vs) {
          const x = el("button", { type: "button", class: "x", text: "×", "aria-label": `Ta bort ${v.name}s röst`, title: "Ta bort rösten" });
          x.addEventListener("click", async () => {
            if (x.dataset.armed !== "true") { x.dataset.armed = "true"; x.textContent = "ta bort?"; setTimeout(() => { x.dataset.armed = ""; x.textContent = "×"; }, 4000); return; }
            try { await adminApi(`api/admin/vote?week=${encodeURIComponent(e.id)}&name=${encodeURIComponent(v.name)}`, { method: "DELETE" }); await refresh(); } catch { x.textContent = "fel"; }
          });
          box.append(el("span", {}, v.name, el("b", { text: String(v.score) }), x));
        }
        li.append(box);
      }
      ul.append(li);
    }
  }

  if (S.key) tryLogin(S.key);
})();
