// Public page: this week's drink, voting (office password, then rating, then name) and the top list
(() => {
  "use strict";
  const { el, $, api, store } = FD;
  const PASS_KEY = "fd-pass", NAME_KEY = "fd-name";
  const S = { data: null, chosen: null, score: null, pass: store.get(PASS_KEY), name: store.get(NAME_KEY) || "", sending: false, msg: null, lockedUntil: null };

  const past = () => (S.data?.weeks || []).slice();
  const current = () => {
    const w = past(); if (!w.length) return null;
    return w.find((e) => e.id === S.chosen) || w.find((e) => e.id === S.data.now.id) || w[w.length - 1];
  };

  function lockText(cur) {
    const f = FD.fridayOf(cur.year, cur.week);
    if (FD.todayInStockholm() === f.toISOString().slice(0, 10)) {
      const mins = Math.ceil((cur.opensAt - Date.now()) / 60000);
      return mins <= 60 ? `Röstningen öppnar i dag kl. 15:40, om ${mins} min.` : "Röstningen öppnar i dag kl. 15:40.";
    }
    return `Röstningen öppnar fredag ${FD.fmtLong(f)} kl. 15:40.`;
  }

  function renderVote(animate) {
    const body = $("vote-body"); body.replaceChildren();
    const now = S.data?.now;
    if (!S.data) { body.append(el("p", { class: "week-line lbl", text: "Veckans drink" }), el("p", { class: "empty-big", text: S.loadError ? "Kunde inte ladda veckans drink. Ladda om sidan." : "Laddar…" })); return; }
    const cur = current();
    if (!cur) { body.append(el("p", { class: "week-line lbl", text: `Vecka ${now.week}` }), el("p", { class: "empty-big", text: "Ingen drink är inlagd än." })); return; }
    FD.setWeekImage(document.querySelector(".vote-img"), cur.week);

    const isNow = cur.id === now.id;
    body.append(el("p", { class: "week-line lbl", text: isNow ? `Veckans drink, vecka ${cur.week}` : `Vecka ${cur.week}${cur.year !== now.year ? " " + cur.year : ""}, ${FD.fmtDate(FD.fridayOf(cur.year, cur.week))}` }));
    body.append(el("h2", { class: "drink-name", text: cur.drink || "Namnlös drink" }));
    if (cur.ingredients) body.append(el("div", { class: "ingredients" }, el("span", { class: "lbl", text: "Ingredienser" }), el("p", { text: cur.ingredients })));

    const open = Date.now() >= cur.opensAt;
    S.lockedUntil = open ? null : cur.opensAt;

    if (!open) {
      body.append(el("p", { class: "status", role: "status", text: lockText(cur) }));
    } else if (!S.pass) {
      // Step 1: office password (remembered on this phone afterwards)
      const input = el("input", { type: "password", id: "pass", autocomplete: "current-password", placeholder: "Kontorets lösenord", "aria-label": "Kontorets lösenord" });
      const go = el("button", { class: "btn lbl", type: "button", text: "Fortsätt" });
      const status = el("p", { class: "status" + (S.msg?.err ? " err" : ""), role: "status", text: S.msg?.text || "Skriv kontorets lösenord för att rösta. Telefonen kommer ihåg det sen." });
      const submit = async () => {
        const v = input.value.trim(); if (!v) { input.focus(); return; }
        go.disabled = true;
        try { await api("api/check", { method: "POST", body: { passcode: v } }); S.pass = v; store.set(PASS_KEY, v); S.msg = null; renderVote(); }
        catch (e) { S.msg = { err: true, text: e.status === 401 ? "Fel lösenord. Fråga Drinkkommittén." : "Något gick fel. Försök igen." }; renderVote(); $("pass")?.focus(); }
      };
      go.addEventListener("click", submit);
      input.addEventListener("keydown", (e) => { if (e.key === "Enter") submit(); });
      body.append(el("div", { class: "step" }, el("label", { class: "step-label lbl", for: "pass", text: "Lösenord" }), el("div", { class: "pass-row" }, input, go)), status);
    } else {
      // Step 2: rating
      const scale = el("div", { class: "scale step" });
      scale.append(el("div", { class: "scale-head" }, el("span", { text: "1 = inte min grej" }), el("span", { text: "10 = mer av den här" })));
      const grid = el("div", { class: "scores", role: "group", "aria-label": `Betyg för ${cur.drink}, 1 till 10` });
      for (let n = 1; n <= 10; n++) grid.append(el("button", { type: "button", text: String(n), "aria-pressed": String(S.score === n), "aria-label": `Ge betyg ${n}`, onclick: () => { S.score = n; S.msg = null; renderVote(); $("name")?.focus(); } }));
      scale.append(grid); body.append(scale);

      // Step 3: name, then send
      const sel = el("select", { class: "name-select", id: "name", onchange: (e) => { S.name = e.target.value; S.msg = null; renderVote(); $("name")?.focus(); } });
      sel.append(el("option", { value: "", text: "Välj ditt namn", selected: !S.name }));
      for (const n of S.data.names) sel.append(el("option", { value: n, text: n, selected: S.name === n }));
      body.append(el("div", { class: "step" }, el("label", { class: "step-label lbl", for: "name", text: "Vem är du?" }), sel));

      const send = el("button", { class: "btn lbl", type: "button", text: S.sending ? "Skickar…" : "Skicka röst", disabled: !S.score || !S.name || S.sending, onclick: () => sendVote(cur) });
      body.append(el("div", { class: "submit-row" }, send));
      body.append(el("p", { class: "status" + (S.msg?.err ? " err" : S.msg ? " ok" : ""), role: "status", text: S.msg?.text || (!S.score ? "Välj ett betyg från 1 till 10." : !S.name ? "Välj ditt namn och skicka." : "Röstar du igen ersätts din förra röst.") }));
    }

    const weeks = past();
    if (weeks.length > 1) {
      const sel = el("select", { "aria-label": "Välj vecka att betygsätta", onchange: (e) => { S.chosen = e.target.value; S.score = null; S.msg = null; renderVote(true); body.querySelector(".other-week select")?.focus(); } });
      for (const e of weeks.slice(-12).reverse()) sel.append(el("option", { value: e.id, text: `v. ${e.week}${e.year !== now.year ? " " + e.year : ""}: ${e.drink}`, selected: e.id === cur.id }));
      body.append(el("div", { class: "other-week" }, el("span", { text: "Missade du en fredag?" }), sel));
    }
    if (S.pass && open) body.append(el("div", { class: "other-week" }, el("button", { type: "button", class: "link", text: "Byt lösenord", onclick: () => { S.pass = null; store.del(PASS_KEY); renderVote(); } })));
    if (animate && !FD.reduceMotion) body.animate([{ opacity: 0, transform: "translateY(6px)" }, { opacity: 1, transform: "none" }], { duration: 380, easing: "cubic-bezier(.2,.7,.2,1)" });
  }

  async function sendVote(cur) {
    if (S.sending) return;
    S.sending = true; renderVote();
    try {
      await api("api/vote", { method: "POST", body: { passcode: S.pass, weekId: cur.id, name: S.name, score: S.score } });
      store.set(NAME_KEY, S.name);
      S.msg = { text: `Tack, ${S.name}! Du gav ${cur.drink} ${S.score}. Röstar du igen ersätts rösten.` };
      await load();
    } catch (e) {
      if (e.code === "bad_pass") { S.pass = null; store.del(PASS_KEY); S.msg = { err: true, text: "Lösenordet stämmer inte längre. Skriv det igen." }; }
      else S.msg = { err: true, text: e.message || "Rösten sparades inte. Försök igen." };
    } finally { S.sending = false; renderVote(); }
  }

  async function load() {
    try { S.data = await api("api/state"); S.loadError = false; }
    catch { S.loadError = true; }
    renderVote(); if (S.data) FD.renderBoard($("board"), S.data.weeks, S.data.now.year);
  }

  load();
  setInterval(() => { if (S.lockedUntil) renderVote(); }, 30000);                     // unlock at 15:40, keep countdown fresh
  setInterval(() => { if (!document.hidden && !S.sending && !document.activeElement?.matches("input,select")) load(); }, 60000); // keep the top list fresh
})();
