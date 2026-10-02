// TV mode: big drink, QR code to the voting page, and a top list that refreshes every 10 seconds
(() => {
  "use strict";
  const { el, $, api } = FD;
  $("tv-qr").innerHTML = FD.qrSvg(FD.voteUrl());

  async function tick() {
    let d;
    try { d = await api("api/state"); } catch { $("tv-foot").textContent = "Tappade kontakten, försöker igen…"; return; }
    const cur = d.weeks.find((w) => w.id === d.now.id) || d.weeks[d.weeks.length - 1];
    const box = $("tv-drink"); box.replaceChildren();
    if (!cur) box.append(el("p", { class: "week-line lbl", text: `Vecka ${d.now.week}` }), el("h1", { class: "drink-name", text: "Ingen drink inlagd" }));
    else {
      box.append(el("p", { class: "week-line lbl", text: cur.id === d.now.id ? `Veckans drink, vecka ${cur.week}` : `Vecka ${cur.week}` }), el("h1", { class: "drink-name", text: cur.drink }));
      if (cur.ingredients) box.append(el("div", { class: "ingredients" }, el("span", { class: "lbl", text: "Ingredienser" }), el("p", { text: cur.ingredients })));
      const open = Date.now() >= cur.opensAt;
      box.append(el("p", { class: "status", text: open ? `${cur.votes} ${cur.votes === 1 ? "röst" : "röster"} hittills` : "Röstningen öppnar kl. 15:40" }));
    }
    FD.renderBoard($("board"), d.weeks, d.now.year, 8);
    $("tv-foot").textContent = `Uppdaterad ${new Date().toLocaleTimeString("sv-SE", { hour: "2-digit", minute: "2-digit" })}`;
  }
  tick();
  setInterval(tick, 10000);
})();
