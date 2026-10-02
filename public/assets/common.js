// Shared browser helpers for Fredagsdrinken pages
(() => {
  "use strict";
  const pad = (n) => String(n).padStart(2, "0");

  const FD = {
    wid: (y, w) => `${y}-W${pad(w)}`,
    $: (id) => document.getElementById(id),
    reduceMotion: matchMedia("(prefers-reduced-motion: reduce)").matches,

    isoWeek(d) {
      const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      const day = t.getUTCDay() || 7;
      t.setUTCDate(t.getUTCDate() + 4 - day);
      const y = t.getUTCFullYear();
      return { year: y, week: Math.ceil(((t - Date.UTC(y, 0, 1)) / 86400000 + 1) / 7) };
    },
    weeksIn(y) { return FD.isoWeek(new Date(y, 11, 28)).week; },
    fridayOf(y, w) {
      const jan4 = new Date(Date.UTC(y, 0, 4));
      const d = new Date(jan4);
      d.setUTCDate(jan4.getUTCDate() - ((jan4.getUTCDay() || 7) - 1) + (w - 1) * 7 + 4);
      return d;
    },
    fmtDate: (d) => d.toLocaleDateString("sv-SE", { weekday: "short", day: "numeric", month: "short", timeZone: "UTC" }),
    fmtLong: (d) => d.toLocaleDateString("sv-SE", { day: "numeric", month: "long", timeZone: "UTC" }),
    fmtAvg: (n) => n.toLocaleString("sv-SE", { minimumFractionDigits: 1, maximumFractionDigits: 1 }),
    todayInStockholm: () => new Date().toLocaleDateString("sv-SE", { timeZone: "Europe/Stockholm" }),

    el(tag, props = {}, ...kids) {
      const n = document.createElement(tag);
      for (const [k, v] of Object.entries(props)) {
        if (k === "class") n.className = v;
        else if (k === "text") n.textContent = v;
        else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
        else if (v !== undefined && v !== null && v !== false) n.setAttribute(k, v === true ? "" : v);
      }
      for (const c of kids) if (c != null) n.append(c);
      return n;
    },

    async api(path, opts = {}) {
      const headers = { ...(opts.body ? { "content-type": "application/json" } : {}), ...(opts.headers || {}) };
      const res = await fetch(path, { ...opts, headers, body: opts.body ? JSON.stringify(opts.body) : undefined });
      let data = null;
      try { data = await res.json(); } catch { /* not json */ }
      if (!res.ok) { const e = new Error(data?.error || `HTTP ${res.status}`); e.status = res.status; e.code = data?.code; throw e; }
      return data;
    },

    store: {
      get(k) { try { return localStorage.getItem(k); } catch { return null; } },
      set(k, v) { try { localStorage.setItem(k, v); } catch { /* private mode */ } },
      del(k) { try { localStorage.removeItem(k); } catch { /* private mode */ } },
    },

    // Ten images rotate by week number: week % 10 picks one, so each drink keeps its picture
    IMGS: Array.from({ length: 10 }, (_, i) => `img/drink-${i}.jpg`),
    setWeekImage(box, week) {
      if (!box) return;
      const n = ((week % FD.IMGS.length) + FD.IMGS.length) % FD.IMGS.length;
      if (box.dataset.img === String(n)) return;
      box.dataset.img = String(n);
      if (!box.children.length) box.append(FD.el("div", { class: "layer" }), FD.el("div", { class: "layer" }));
      const [a, b] = box.children;
      const next = a.classList.contains("on") ? b : a, prev = next === a ? b : a;
      const src = FD.IMGS[n];
      const show = () => { if (box.dataset.img !== String(n)) return; next.style.backgroundImage = `url("${src}")`; next.classList.add("on"); prev.classList.remove("on"); };
      const img = new Image(); img.src = src;
      (img.decode ? img.decode() : Promise.resolve()).then(show, show);
    },

    // Leaderboard: group weeks by drink name, average all votes, round to one decimal
    leaderboard(weeks, nowYear) {
      const groups = new Map();
      for (const e of weeks) {
        if (!e.votes) continue;
        const key = (e.drink || "").trim().toLowerCase();
        if (!key) continue;
        const g = groups.get(key) || { name: e.drink.trim(), sum: 0, n: 0, weeks: [] };
        g.sum += e.sum; g.n += e.votes; g.weeks.push(e);
        groups.set(key, g);
      }
      const list = [...groups.values()].map((g) => ({ ...g, avg: Math.round((g.sum / g.n) * 10) / 10 }))
        .sort((a, b) => b.avg - a.avg || b.n - a.n || a.name.localeCompare(b.name, "sv"));
      let prev = null, rank = 0;
      list.forEach((g, i) => { if (g.avg !== prev) rank = i + 1; prev = g.avg; g.rank = rank; g.wk = g.weeks.map((w) => `v. ${w.week}${w.year !== nowYear ? " " + w.year : ""}`).join(", "); });
      return list;
    },
    renderBoard(box, weeks, nowYear, limit = Infinity) {
      const { el } = FD;
      box.replaceChildren();
      const list = FD.leaderboard(weeks, nowYear).slice(0, limit);
      if (!list.length) { box.append(el("p", { class: "board-empty", text: "Topplistan fylls på när den första drinken har fått betyg." })); return; }
      const ol = el("ol", { class: "rows" });
      for (const g of list) {
        ol.append(el("li", { class: "row" + (g.rank === 1 ? " first" : "") },
          el("span", { class: "rank", text: String(g.rank), "aria-label": `Plats ${g.rank}` }),
          el("span", { class: "name", text: g.name }),
          el("span", { class: "avg", text: FD.fmtAvg(g.avg), "aria-label": `Snittbetyg ${FD.fmtAvg(g.avg)} av 10` }),
          el("span", { class: "meta", text: `${g.n} ${g.n === 1 ? "röst" : "röster"}, serverad ${g.wk}` })));
      }
      box.append(ol);
    },

    // QR code as inline SVG (uses the vendored qrcode-generator library)
    qrSvg(text) {
      const qr = window.qrcode(0, "M");
      qr.addData(text); qr.make();
      return qr.createSvgTag({ cellSize: 4, margin: 2, scalable: true });
    },
    voteUrl: () => new URL("./#rosta", location.href).href,
  };
  window.FD = FD;
})();
