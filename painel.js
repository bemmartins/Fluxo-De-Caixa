/* =====================================================================
   Fluxo de Caixa – Tesouraria · painel
   Regras de negócio, telas, gráficos, relatório para imprimir, pasta do
   computador e conexão com o SharePoint (Microsoft Graph).
   Veja o resumo das regras no comentário do topo de index.html.
   Para deixar o SharePoint configurado para todos: procure SP_DEFAULT.
   ===================================================================== */

(function () {
  /* Volta do login da Microsoft (janela pop-up): entrega o código para a janela do painel e fecha */
  {
    const q = new URLSearchParams(location.search);
    if (q.get("state") && (q.get("code") || q.get("error"))) {
      const ret = {
        fcAuth: true,
        code: q.get("code"),
        error: q.get("error"),
        desc: q.get("error_description"),
        state: q.get("state"),
      };
      try {
        localStorage.setItem("fc-oauth-ret", JSON.stringify(ret));
      } catch (e) {}
      try {
        if (window.opener) window.opener.postMessage(ret, location.origin);
      } catch (e) {}
      document.body.innerHTML =
        '<p style="font:16px system-ui,sans-serif;color:#EAF3FB;padding:28px">Login concluído. Esta janela vai fechar.</p>';
      setTimeout(() => {
        try {
          window.close();
        } catch (e) {}
      }, 400);
      return;
    }
  }
  /* As funções de leitura das planilhas (parseBook, planoFromName, norm…) ficam em js/leitor.js,
   carregado antes deste arquivo. */

  const LOGO = document.querySelector(".logo").getAttribute("src");
  const FRAMED = (() => {
    try {
      return window.self !== window.top;
    } catch (e) {
      return true;
    }
  })();
  const $ = (s) => document.querySelector(s);
  const STORE = "fluxo-caixa-tesouraria-v3";
  const MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  const MESES_L = [
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
  ];
  const DIAS_L = ["domingo", "segunda", "terça", "quarta", "quinta", "sexta", "sábado"];
  const BRL = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  const INT = new Intl.NumberFormat("pt-BR");
  const esc = (s) =>
    String(s ?? "").replace(
      /[&<>"']/g,
      (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c],
    );
  const toD = (k) => {
    const [y, m, d] = k.split("-").map(Number);
    return new Date(y, m - 1, d, 12);
  };
  const addDays = (k, n) => {
    const d = toD(k);
    d.setDate(d.getDate() + n);
    return fromD(d);
  };
  const diffDays = (a, b) => Math.round((toD(b) - toD(a)) / 864e5);
  const fDay = (k) => {
    const [y, m, d] = k.split("-");
    return `${d}/${m}/${y.slice(2)}`;
  };
  const fDayL = (k) => {
    const [y, m, d] = k.split("-");
    return `${d}/${m}/${y}`;
  };
  const fDM = (k) => {
    const [, m, d] = k.split("-");
    return `${+d} ${MESES[+m - 1]}`;
  };
  const money = (v) => BRL.format(Math.abs(v) < 0.005 ? 0 : v);
  const big = (v, signed) =>
    Math.abs(v) >= 1e8
      ? (signed ? (v < 0 ? "− " : "+ ") : "") + short(Math.abs(v)).replace("−", "")
      : signed
        ? sgn(v)
        : money(v);
  const sgn = (v) => {
    if (Math.abs(v) < 0.005) v = 0;
    return `${v < 0 ? "−" : "+"} ${money(Math.abs(v))}`;
  };
  function short(v) {
    const a = Math.abs(v),
      s = v < 0 ? "−" : "";
    let t;
    if (a >= 1e9) t = (a / 1e9).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + " bi";
    else if (a >= 1e6) t = (a / 1e6).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + " mi";
    else if (a >= 1e3) t = (a / 1e3).toLocaleString("pt-BR", { maximumFractionDigits: 0 }) + " mil";
    else t = a.toLocaleString("pt-BR", { maximumFractionDigits: 0 });
    return s + "R$ " + t;
  }
  const pct = (a, b) => {
    if (!b) return "—";
    let v = (a / b) * 100;
    if (a < b && v >= 99.5)
      return (Math.floor(v * 10) / 10).toLocaleString("pt-BR", { maximumFractionDigits: 1 }) + "%";
    return v.toLocaleString("pt-BR", { maximumFractionDigits: v < 10 ? 1 : 0 }) + "%";
  };
  const lastDayOf = (k) => {
    const d = toD(k);
    d.setMonth(d.getMonth() + 1, 0);
    return fromD(d);
  };
  const stamp = (d) => {
    d = d instanceof Date ? d : new Date(d);
    const same = d.toDateString() === new Date().toDateString();
    return `${same ? "hoje" : pad(d.getDate()) + "/" + pad(d.getMonth() + 1)} às ${pad(d.getHours())}h${pad(d.getMinutes())}`;
  };
  const plural = (n, a, b) => `${INT.format(n)} ${n === 1 ? a : b}`;

  /* ---------- planos ---------- */
  const ICON = {
    taxa: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M19 5 5 19"/><circle cx="7" cy="7" r="2.5"/><circle cx="17" cy="17" r="2.5"/></svg>',
    financeiro:
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6"/></svg>',
    militar:
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3 5 6v5c0 4.5 3 8.3 7 10 4-1.7 7-5.5 7-10V6z"/><path d="m9 12 2 2 4-4"/></svg>',
    previdenciario:
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 0 1 18 0z"/><path d="M12 12v6a2.5 2.5 0 0 1-5 0"/></svg>',
    sem: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M9.5 9.5a2.5 2.5 0 1 1 3.5 2.3c-.6.3-1 .9-1 1.6M12 17h.01"/></svg>',
  };
  const PL = [
    { k: "taxa", l: "Taxa", full: "Taxa de Administração", c: "var(--taxa)" },
    { k: "financeiro", l: "Financeiro", full: "Plano Financeiro", c: "var(--fin)" },
    { k: "militar", l: "Militar", full: "Plano Militar", c: "var(--mil)" },
    { k: "previdenciario", l: "Previdenciário", full: "Plano Previdenciário", c: "var(--prev)" },
  ];
  const PLM = Object.fromEntries(PL.map((p) => [p.k, p]));
  const SEM = { k: "sem", l: "Sem plano", full: "Planilhas sem plano definido", c: "var(--sem)" };
  const plInfo = (k) => PLM[k] || SEM;

  /* ---------- dados ---------- */
  let state = { accounts: {}, updatedAt: null, embedded: false };
  let booting = true;
  let ALL = [],
    ORDER = [];
  const ui = {
    view: "geral",
    preset: "mes",
    from: null,
    to: null,
    open: new Set(),
    q: "",
    limit: 50,
    rplan: "todos",
    rsub: false,
  };
  const BOOKS = {};

  /* Plano da planilha: escolha manual > nome/pasta/conteúdo reconhecidos. */
  function planoOf(acc) {
    const o = (state.planoSet || {})[acc.srcName || acc.name];
    if (o) return o;
    return (
      acc.planoHint ||
      planoFromName((acc.paths || []).join(" ") + " " + (acc.file || "") + " " + acc.name) ||
      ""
    );
  }
  /* Natureza usada nos agrupamentos ("Não Encontrado" vira "Natureza não identificada"). */
  function natLabel(t) {
    const raw = t.tipo === "DESPESA" ? t.natD || t.nat : t.nat;
    if (!raw || norm(raw) === "nao encontrado") return "Natureza não identificada";
    return raw;
  }
  /* ---------- aportes ----------
   NP = aporte (positiva entra, negativa sai). Das transferências, só entram as do Financeiro
   para as contas do Militar: saem do Financeiro como repasse e entram no Militar como aporte recebido. */
  const APO_FIN_MIL = "Repasse de aporte ao Plano Militar",
    APO_MIL_FIN = "Aporte recebido do Plano Financeiro";
  /* Natureza de uma linha NP (aporte). */
  function aporteNat(t) {
    if (t.value < 0) return "Saída de aporte (NP)";
    const raw = t.nat || t.natD;
    return raw && norm(raw) !== "nao encontrado" ? raw : "Aporte (NP)";
  }
  const digits = (s) => String(s || "").replace(/\D/g, "");
  const destKey = (c) => norm(c).replace(/\s+/g, "");
  /* contas do Militar: pelo nome das planilhas do plano Militar (ex.: "BB - 292.061" → 292061) e pelas digitadas na aba Planilhas */
  /* Números das contas do Militar: tirados do nome das planilhas do Militar + digitados na aba Planilhas. */
  function militarNums() {
    const auto = new Set(),
      man = new Set();
    for (const nm of ORDER) {
      const a = state.accounts[nm];
      if ((planoOf(a) || "") !== "militar") continue;
      const src = String(a.srcName || a.name);
      const parts = src.split(/\s+[-–]\s+/);
      const d = digits(parts.length > 1 ? parts.slice(1).join(" ") : src).replace(/^0+/, "");
      if (d.length >= 4) auto.add(d);
    }
    for (const n of state.aporteNums || []) {
      const d = digits(n).replace(/^0+/, "");
      if (d.length >= 4) man.add(d);
    }
    return { auto: [...auto], man: [...man], all: [...new Set([...auto, ...man])] };
  }
  /* Diz se a conta de destino de uma transferência parece ser uma conta do Militar. */
  function destSuggested(credor, nums) {
    const segs = String(credor || "")
      .split(/\s*[-–]\s*/)
      .map((x) => digits(x).replace(/^0+/, ""))
      .filter(Boolean);
    const last = segs[segs.length - 1] || "";
    return !!last && nums.some((n) => last === n || (last.startsWith(n) && last.length - n.length <= 1));
  }
  /* Transferência do Financeiro é aporte para o Militar? (escolha manual ou sugestão). */
  function isAporteDest(credor, nums) {
    const ex = (state.aporteDest || {})[destKey(credor)];
    return ex != null ? ex : destSuggested(credor, nums);
  }
  const subKey = (t) =>
    t.tipo === "RECEITA"
      ? t.ident || t.idk || t.item || t.lanc
      : t.tipo === "APORTE"
        ? t.mirror
          ? t.natK === APO_MIL_FIN
            ? `Do Financeiro · ${t.srcAcct}`
            : `Para o Militar · ${t.srcAcct}`
          : t.natK === APO_MIL_FIN
            ? `Da conta ${t.credor}`
            : t.natK === APO_FIN_MIL
              ? `Para a conta ${t.credor}`
              : splitCode(t.credor).name || t.ident || t.lanc
        : splitCode(t.credor).name || t.item || t.lanc;
  const c2 = (v) => Math.round(v * 100) / 100 || 0;
  const apo = (L) => {
    let s = 0;
    for (const t of L) if (t.tipo === "APORTE") s += t.value;
    return c2(s);
  };
  const apoIn = (L) => {
    let s = 0;
    for (const t of L) if (t.tipo === "APORTE" && t.value > 0) s += t.value;
    return c2(s);
  };
  const apoOut = (L) => {
    let s = 0;
    for (const t of L) if (t.tipo === "APORTE" && t.value < 0) s -= t.value;
    return c2(s);
  };
  const resOf = (L) => rec(L) + apo(L) - des(L);
  /* Monta a lista ALL com todos os lançamentos que entram no painel (receita, despesa e aporte),
   aplicando as regras de aporte e criando o lado espelhado Financeiro ↔ Militar. */
  function rebuild() {
    ORDER = Object.keys(state.accounts).sort((a, b) => a.localeCompare(b, "pt-BR"));
    ALL = [];
    let id = 0;
    const nums = militarNums().all;
    const repasses = [],
      recebidos = [];
    const srch = (t, nat) =>
      norm(
        [
          t.tipo === "APORTE" ? "aporte" : "",
          t.lanc,
          t.ident,
          t.idk,
          nat,
          t.credor,
          t.item,
          t.doc,
          t.ob,
          t.proc,
        ].join(" "),
      );
    for (const nm of ORDER) {
      const acc = state.accounts[nm];
      const p = planoOf(acc) || "sem";
      for (const t of acc.txs) {
        let tipo = t.tipo,
          nat;
        if (tipo === "NP") {
          tipo = "APORTE";
          nat = aporteNat(t);
        } else if (tipo === "TRANSF") {
          if (p === "financeiro" && t.value < 0 && t.credor && isAporteDest(t.credor, nums)) {
            tipo = "APORTE";
            nat = APO_FIN_MIL;
          } else if (
            p === "militar" &&
            t.value > 0 &&
            t.credor &&
            (state.aporteOrig || {})[destKey(t.credor)]
          ) {
            tipo = "APORTE";
            nat = APO_MIL_FIN;
          } else continue;
        } else if (tipo === "RECEITA" || tipo === "DESPESA") nat = natLabel(t);
        else continue;
        const row = { ...t, tipo, acct: nm, plano: p, natK: nat, id: id++ };
        row.s = srch(row, nat);
        ALL.push(row);
        if (nat === APO_FIN_MIL) repasses.push(row);
        else if (nat === APO_MIL_FIN) recebidos.push(row);
      }
    }
    /* o mesmo valor entra no Militar como aporte recebido (sem duplicar se a planilha do Militar já tiver a NP) */
    const milAp = ALL.filter((t) => t.plano === "militar" && t.tipo === "APORTE");
    for (const t of repasses) {
      const v = -t.value;
      if (milAp.some((x) => x.date === t.date && Math.abs(x.value - v) < 0.01)) continue;
      const row = {
        ...t,
        value: v,
        plano: "militar",
        natK: APO_MIL_FIN,
        mirror: true,
        srcAcct: t.acct,
        id: id++,
      };
      row.s = srch(row, APO_MIL_FIN) + " financeiro";
      ALL.push(row);
    }
    /* entrada marcada no Militar sem a saída correspondente no Financeiro: o Financeiro mostra o repasse */
    for (const t of recebidos) {
      if (repasses.some((x) => x.date === t.date && Math.abs(-x.value - t.value) < 0.01)) continue;
      const row = {
        ...t,
        value: -t.value,
        plano: "financeiro",
        natK: APO_FIN_MIL,
        mirror: true,
        srcAcct: t.acct,
        id: id++,
      };
      row.s = srch(row, APO_FIN_MIL) + " militar";
      ALL.push(row);
    }
    renderBadge();
  }

  /* ---------- armazenamento ---------- */
  const IDB = {
    db: null,
    open() {
      if (this.db) return Promise.resolve(this.db);
      return new Promise((res, rej) => {
        let r;
        try {
          r = indexedDB.open("fluxo-caixa-tesouraria", 1);
        } catch (e) {
          rej(e);
          return;
        }
        r.onupgradeneeded = () => r.result.createObjectStore("kv");
        r.onsuccess = () => {
          this.db = r.result;
          res(r.result);
        };
        r.onerror = () => rej(r.error);
        r.onblocked = () => rej(new Error("bloqueado"));
      });
    },
    async get(k) {
      const db = await this.open();
      return new Promise((res, rej) => {
        const q = db.transaction("kv", "readonly").objectStore("kv").get(k);
        q.onsuccess = () => res(q.result);
        q.onerror = () => rej(q.error);
      });
    },
    async set(k, v) {
      const db = await this.open();
      return new Promise((res, rej) => {
        const tx = db.transaction("kv", "readwrite");
        tx.objectStore("kv").put(v, k);
        tx.oncomplete = () => res();
        tx.onerror = () => rej(tx.error);
        tx.onabort = () => rej(tx.error);
      });
    },
  };
  /* Guarda as planilhas e as escolhas no navegador (IndexedDB; localStorage de reserva). */
  async function save() {
    if (state.embedded) return true;
    const manual = {};
    for (const [k, a] of Object.entries(state.accounts)) if (a.src !== "pasta") manual[k] = a;
    const data = {
      accounts: manual,
      files: state.files || {},
      srcKind: state.srcKind || "",
      folderName: state.folderName || "",
      folderRead: state.folderRead || null,
      alias: state.alias || {},
      tabPref: state.tabPref || {},
      planoSet: state.planoSet || {},
      aporteDest: state.aporteDest || {},
      aporteOrig: state.aporteOrig || {},
      aporteNums: state.aporteNums || [],
      updatedAt: state.updatedAt,
      embedded: false,
    };
    try {
      await IDB.set("state", data);
      try {
        localStorage.removeItem(STORE);
      } catch (e) {}
      return true;
    } catch (e) {
      try {
        localStorage.setItem(STORE, JSON.stringify(data));
        return true;
      } catch (e2) {
        return false;
      }
    }
  }
  /* Carrega o que foi guardado no navegador. */
  async function load() {
    const ok = (s) =>
      s && s.accounts && (Object.keys(s.accounts).length || Object.keys(s.files || {}).length);
    try {
      const s = await IDB.get("state");
      if (ok(s)) {
        state = s;
        compose();
        return true;
      }
    } catch (e) {}
    try {
      const s = JSON.parse(localStorage.getItem(STORE) || "null");
      if (ok(s)) {
        state = s;
        compose();
        return true;
      }
    } catch (e) {}
    return false;
  }
  function savePrefs() {
    try {
      localStorage.setItem(
        STORE + "-prefs2",
        JSON.stringify({ preset: ui.preset, rplan: ui.rplan, rsub: ui.rsub }),
      );
    } catch (e) {}
  }

  /* ---------- período ---------- */
  function bounds() {
    let mn = null,
      mx = null;
    for (const t of ALL) {
      if (!mn || t.date < mn) mn = t.date;
      if (!mx || t.date > mx) mx = t.date;
    }
    return { mn, mx };
  }
  /* Período escolhido (Dia, Semana, Mês, Ano, Tudo ou datas) a partir do último dia com dados. */
  function range() {
    const { mn, mx } = bounds();
    if (!mx) return {};
    let from,
      to = mx;
    if (ui.preset === "custom") {
      from = ui.from || mn;
      to = ui.to || mx;
      if (from > to) [from, to] = [to, from];
    } else if (ui.preset === "all") from = mn;
    else if (ui.preset === "dia") from = mx;
    else if (ui.preset === "semana") from = addDays(mx, -6);
    else if (ui.preset === "mes") from = mx.slice(0, 8) + "01";
    else if (ui.preset === "ano") from = mx.slice(0, 5) + "01-01";
    return { from, to };
  }
  /* Período anterior equivalente, para as comparações (▲▼). */
  function prevRange(from, to) {
    if (ui.preset === "mes") {
      const d = toD(from);
      d.setMonth(d.getMonth() - 1);
      const pf = fromD(d);
      const pe = addDays(from, -1);
      return { pf, pt: pf.slice(0, 8) + pad(Math.min(+to.slice(8), +pe.slice(8))), lab: "mês anterior" };
    }
    if (ui.preset === "ano") {
      const y = +from.slice(0, 4) - 1;
      return { pf: `${y}-01-01`, pt: `${y}${to.slice(4)}`, lab: "ano anterior" };
    }
    if (ui.preset === "dia") {
      const prev = ALL.filter((t) => t.date < from).reduce((m, t) => (t.date > m ? t.date : m), "");
      return prev ? { pf: prev, pt: prev, lab: `dia ${fDay(prev)}` } : null;
    }
    const span = diffDays(from, to) + 1;
    return { pf: addDays(from, -span), pt: addDays(from, -1), lab: `${span} dias anteriores` };
  }
  function periodTitle(from, to) {
    const M = (k) => MESES_L[+k.slice(5, 7) - 1];
    if (ui.preset === "dia") return `${DIAS_L[toD(to).getDay()]}, ${fDayL(to)}`;
    if (ui.preset === "semana") return `${fDM(from)} a ${fDM(to)}`;
    if (ui.preset === "mes")
      return `${M(to)} de ${to.slice(0, 4)}${to < lastDayOf(to) ? ` até dia ${+to.slice(8)}` : ""}`;
    if (ui.preset === "ano") return `${to.slice(0, 4)} até ${fDM(to)}`;
    return `${fDayL(from)} a ${fDayL(to)}`;
  }
  const PRESETS = [
    ["dia", "Dia"],
    ["semana", "Semana"],
    ["mes", "Mês"],
    ["ano", "Ano"],
    ["all", "Tudo"],
  ];
  function periodCtrl(from, to) {
    return `<div class="ctrl"><div class="seg" role="group" aria-label="Período">${PRESETS.map(([k, l]) => `<button type="button" data-preset="${k}" aria-pressed="${ui.preset === k}">${l}</button>`).join("")}</div>
    <label class="dates">de <input type="date" id="dFrom" value="${from}"> até <input type="date" id="dTo" value="${to}"></label></div>`;
  }

  /* ---------- somas ---------- */
  const rec = (L) => {
    let s = 0;
    for (const t of L) if (t.tipo === "RECEITA") s += t.value;
    return s;
  };
  const des = (L) => {
    let s = 0;
    for (const t of L) if (t.tipo === "DESPESA") s -= t.value;
    return s;
  };
  /* Lançamentos do período e do período anterior. */
  function ctx() {
    const { from, to } = range();
    const inP = ALL.filter((t) => t.date >= from && t.date <= to);
    const pr = prevRange(from, to);
    const prev = pr ? ALL.filter((t) => t.date >= pr.pf && t.date <= pr.pt) : [];
    return { from, to, inP, pr, prev };
  }
  function deltaChip(cur, prev, goodUp, lab) {
    if (!prev) return `<span class="chip neu">sem comparação</span>`;
    const p = ((cur - prev) / Math.abs(prev)) * 100;
    if (!isFinite(p)) return "";
    const up = p >= 0,
      good = goodUp ? up : !up;
    return `<span class="chip ${good ? "" : "down"}" title="em relação ao ${esc(lab || "período anterior")}">${up ? "▲" : "▼"} ${Math.abs(p).toLocaleString("pt-BR", { maximumFractionDigits: 1 })}% <span style="font-weight:400;opacity:.8">vs ${esc(lab || "anterior")}</span></span>`;
  }

  /* ---------- gráfico de área suave ---------- */
  function niceStep(raw) {
    const p = Math.pow(10, Math.floor(Math.log10(raw)));
    const f = raw / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
  }
  function ticks(lo, hi, n) {
    if (hi - lo <= 0) hi = lo + 1;
    const st = niceStep((hi - lo) / n);
    const a = Math.floor(lo / st + 1e-9) * st,
      b = Math.ceil(hi / st - 1e-9) * st;
    const out = [];
    for (let v = a; v <= b + st / 2; v += st) out.push(Math.round(v / st) * st);
    return out;
  }
  function bucketOf(k, g) {
    if (g === "d") return k;
    if (g === "m") return k.slice(0, 8) + "01";
    const d = toD(k);
    d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
    return fromD(d);
  }
  /* Agrupa os lançamentos por dia, semana ou mês, conforme o tamanho do período. */
  function buckets(rows, from, to) {
    let span = diffDays(from, to) + 1;
    let f = from;
    if (span < 14) {
      f = addDays(to, -13);
      span = 14;
    }
    const g = span <= 45 ? "d" : span <= 200 ? "w" : "m";
    const list = [],
      map = {};
    for (let k = f; k <= to; k = addDays(k, 1)) {
      const b = bucketOf(k, g);
      if (!map[b]) {
        map[b] = { key: b, start: k, end: k, r: 0, a: 0, d: 0 };
        list.push(map[b]);
      }
      map[b].end = k;
    }
    for (const t of rows) {
      const b = map[bucketOf(t.date, g)];
      if (!b) continue;
      if (t.tipo === "RECEITA") b.r += t.value;
      else if (t.tipo === "APORTE") b.a += t.value;
      else b.d -= t.value;
    }
    return { list, g, from: f };
  }
  /* Transforma pontos em curva suave (SVG), sem passar abaixo da linha do zero. */
  function smooth(pts, base) {
    if (pts.length < 2) return pts.length ? `M${pts[0][0]} ${pts[0][1]}` : "";
    let d = `M${pts[0][0].toFixed(1)} ${pts[0][1].toFixed(1)}`;
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i - 1] || pts[i],
        p1 = pts[i],
        p2 = pts[i + 1],
        p3 = pts[i + 2] || p2;
      const t = 0.18;
      const c1 = [p1[0] + (p2[0] - p0[0]) * t, p1[1] + (p2[1] - p0[1]) * t],
        c2 = [p2[0] - (p3[0] - p1[0]) * t, p2[1] - (p3[1] - p1[1]) * t];
      if (base != null) {
        c1[1] = Math.min(c1[1], base);
        c2[1] = Math.min(c2[1], base);
      }
      d += `C${c1[0].toFixed(1)} ${c1[1].toFixed(1)} ${c2[0].toFixed(1)} ${c2[1].toFixed(1)} ${p2[0].toFixed(1)} ${p2[1].toFixed(1)}`;
    }
    return d;
  }
  const bLabel = (b, g) => {
    if (g === "d") return fDM(b.key);
    if (g === "w") return fDM(b.start);
    const [y, m] = b.key.split("-");
    return `${MESES[+m - 1]}/${y.slice(2)}`;
  };
  const bTitle = (b, g) => {
    if (g === "d") return `${DIAS_L[toD(b.key).getDay()]}, ${fDayL(b.key)}`;
    if (g === "w") return `Semana de ${fDM(b.start)} a ${fDM(b.end)}`;
    const [y, m] = b.key.split("-");
    return `${MESES_L[+m - 1]} de ${y}`;
  };
  let chartArgs = null;
  /* Desenha o gráfico Receitas × Aportes × Despesas, com tooltip ao passar o mouse. */
  function drawChart(el, rows, from, to, color) {
    chartArgs = [el, rows, from, to, color];
    const W = Math.max(300, Math.round(el.clientWidth || 760)),
      H = W < 560 ? 220 : 270;
    const { list, g } = buckets(rows, from, to);
    const n = list.length;
    const L = W < 500 ? 50 : 62,
      R = 12,
      T = 14,
      B = 26,
      iw = W - L - R,
      ih = H - T - B;
    const hasA = list.some((b) => b.a > 0);
    const hi = Math.max(1, ...list.map((b) => Math.max(b.r, b.d, b.a)));
    const tk = ticks(0, hi, W < 500 ? 3 : 4),
      y1 = tk[tk.length - 1];
    const X = (i) => L + (n < 2 ? iw / 2 : (i * iw) / (n - 1)),
      Y = (v) => T + (1 - v / y1) * ih;
    const pr = list.map((b, i) => [X(i), Y(b.r)]),
      pd = list.map((b, i) => [X(i), Y(b.d)]),
      pa = list.map((b, i) => [X(i), Y(Math.max(0, b.a))]);
    const id = "g" + Math.random().toString(36).slice(2, 7);
    let s = `<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img" aria-label="Receitas e despesas por ${g === "d" ? "dia" : g === "w" ? "semana" : "mês"}"><defs><linearGradient id="${id}" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${color}" stop-opacity=".42"/><stop offset="1" stop-color="${color}" stop-opacity="0"/></linearGradient></defs>`;
    for (const v of tk)
      s += `<line class="${v === 0 ? "zero" : "grid"}" x1="${L}" x2="${W - R}" y1="${Y(v)}" y2="${Y(v)}"/><text class="ax" x="${L - 8}" y="${Y(v) + 4}" text-anchor="end">${short(v)}</text>`;
    const lr = smooth(pr, Y(0)),
      ld = smooth(pd, Y(0));
    s += `<path d="${lr}L${X(n - 1)} ${Y(0)}L${X(0)} ${Y(0)}Z" fill="url(#${id})"/>`;
    s += `<path d="${ld}" fill="none" stroke="var(--neg)" stroke-width="1.9" stroke-dasharray="1 5" stroke-linecap="round"/>`;
    if (hasA)
      s += `<path class="draw" pathLength="1" d="${smooth(pa, Y(0))}" fill="none" stroke="var(--apo)" stroke-width="1.8" stroke-linecap="round" style="animation-delay:.15s"/>`;
    s += `<path class="draw" pathLength="1" d="${lr}" fill="none" stroke="${color}" stroke-width="2.4" stroke-linecap="round"/>`;
    const step = Math.max(1, Math.ceil(n / Math.max(2, Math.floor(iw / 80))));
    list.forEach((b, i) => {
      if (i % step && i !== n - 1) return;
      if (i !== n - 1 && n - 1 - i < step / 1.6 && i % step === 0 && i !== 0) return;
      s += `<text class="ax" x="${X(i)}" y="${H - 6}" text-anchor="${i === 0 ? "start" : i === n - 1 ? "end" : "middle"}">${bLabel(b, g)}</text>`;
    });
    s += `<g id="hov" visibility="hidden"><line class="guide" x1="0" x2="0" y1="${T}" y2="${T + ih}"/><circle r="5" fill="${color}" stroke="#06131F" stroke-width="2"/><circle r="4" fill="var(--neg)" stroke="#06131F" stroke-width="2"/></g></svg><div class="tip" hidden></div>`;
    el.innerHTML = s;
    const svg = el.querySelector("svg"),
      hov = svg.querySelector("#hov"),
      tip = el.querySelector(".tip");
    const [c1, c2] = hov.querySelectorAll("circle");
    const gl = hov.querySelector("line");
    const move = (e) => {
      const rc = svg.getBoundingClientRect();
      const sx = (e.clientX - rc.left) * (W / rc.width);
      let i = Math.round((sx - L) / (n < 2 ? 1 : iw / (n - 1)));
      i = Math.max(0, Math.min(n - 1, i));
      const b = list[i];
      hov.setAttribute("visibility", "visible");
      gl.setAttribute("x1", X(i));
      gl.setAttribute("x2", X(i));
      c1.setAttribute("cx", X(i));
      c1.setAttribute("cy", Y(b.r));
      c2.setAttribute("cx", X(i));
      c2.setAttribute("cy", Y(b.d));
      tip.innerHTML = `<div class="t">${bTitle(b, g)}</div><div class="r"><span>Receitas</span><span>${money(b.r)}</span></div>${b.a ? `<div class="r"><span>Aportes</span><span>${sgn(b.a)}</span></div>` : ""}<div class="r"><span>Despesas</span><span>${money(b.d)}</span></div><div class="r"><span>Resultado</span><span>${sgn(b.r + b.a - b.d)}</span></div>`;
      tip.hidden = false;
      const sc = rc.width / W;
      let left = X(i) * sc + 14;
      const tw = tip.offsetWidth;
      if (left + tw > rc.width) left = X(i) * sc - tw - 14;
      tip.style.left = Math.max(0, left) + "px";
      tip.style.top = Math.max(0, Math.min(Y(b.r), Y(b.d)) * sc - 56) + "px";
    };
    svg.addEventListener("pointermove", move);
    svg.addEventListener("pointerdown", move);
    svg.addEventListener("pointerleave", () => {
      hov.setAttribute("visibility", "hidden");
      tip.hidden = true;
    });
  }
  /* Minigráfico dos cartões dos planos. */
  function spark(rows, from, to, color) {
    const { list } = buckets(rows, from, to);
    if (list.length < 2) return "";
    const vals = list.map((b) => Math.max(0, b.r + b.a));
    const mx = Math.max(1, ...vals);
    const W = 200,
      H = 34;
    const pts = vals.map((v, i) => [(i / (vals.length - 1)) * W, 3 + (1 - v / mx) * (H - 6)]);
    const d = smooth(pts, H - 3);
    return `<svg class="spark" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" aria-hidden="true"><path d="${d}L${W} ${H}L0 ${H}Z" fill="${color}" opacity=".14"/><path d="${d}" fill="none" stroke="${color}" stroke-width="1.7" vector-effect="non-scaling-stroke"/></svg>`;
  }

  /* ---------- blocos ---------- */
  /* Barra colorida de participação (ex.: receita de cada plano). */
  function shareBar(items, tot) {
    if (!tot) return "";
    return `<div class="share"><div class="segbar">${items
      .filter((x) => x.v > 0)
      .map(
        (x, i) =>
          `<i style="flex:${x.v} 1 0;background:${x.c};animation-delay:${i * 70}ms" title="${esc(x.l)}: ${money(x.v)}"></i>`,
      )
      .join("")}</div>
    <div class="lg">${items.map((x) => `<span><i style="background:${x.c}"></i><em>${esc(x.l)}</em><b>${pct(x.v, tot)}</b></span>`).join("")}</div></div>`;
  }
  /* Barras de resultado mês a mês (últimos 12 meses). */
  function monthBars(rows, to) {
    const end = to.slice(0, 7);
    const ms = [];
    const d = toD(end + "-01");
    for (let i = 11; i >= 0; i--) {
      const x = new Date(d);
      x.setMonth(x.getMonth() - i);
      ms.push(fromD(x).slice(0, 7));
    }
    const by = {};
    for (const m of ms) by[m] = 0;
    for (const t of rows) {
      const m = t.date.slice(0, 7);
      if (m in by) by[m] += t.tipo === "RECEITA" ? t.value : t.value;
    }
    const has = ms.filter((m) => rows.some((t) => t.date.startsWith(m)));
    if (!has.length) return '<p class="empty-note">Sem lançamentos nos últimos 12 meses.</p>';
    const use = ms.slice(ms.indexOf(has[0]));
    const mx = Math.max(0, ...use.map((m) => by[m])),
      mn = Math.min(0, ...use.map((m) => by[m]));
    const span = mx - mn || 1;
    const zp = mx / span; /* fração da altura acima do zero */
    return `<div class="barsbox"><div class="bars">${mn < 0 && mx > 0 ? `<div class="zl" style="top:${(zp * 100).toFixed(2)}%"></div>` : ""}${use
      .map((m, i) => {
        const v = by[m];
        const up = v > 0 && mx > 0 ? (v / mx) * 100 : 0,
          dn = v < 0 && mn < 0 ? (v / mn) * 100 : 0;
        return `<div class="b ${m === end ? "cur" : ""}" title="${MESES_L[+m.slice(5) - 1]}/${m.slice(0, 4)}: ${sgn(v)}"><div style="height:${(zp * 100).toFixed(2)}%;display:flex;flex-direction:column;justify-content:flex-end"><div class="up" style="height:${up.toFixed(2)}%;animation-delay:${i * 35}ms"></div></div><div style="height:${((1 - zp) * 100).toFixed(2)}%"><div class="dn" style="height:${dn.toFixed(2)}%;animation-delay:${i * 35}ms"></div></div><span>${MESES[+m.slice(5) - 1]}</span></div>`;
      })
      .join("")}</div></div>`;
  }
  const CHEV =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m9 6 6 6-6 6"/></svg>';
  function splitCode(s) {
    const m = String(s || "").match(/^([\w.\/]*\d[\w.\/]*)\s+-\s+(.+)$/);
    return m ? { code: m[1], name: m[2] } : { code: "", name: String(s || "") };
  }
  /* ranking por natureza com subdivisões; byPlan = barras empilhadas por plano */
  /* Ranking por natureza com subdivisões (por plano, identificação ou credor) ao clicar. */
  function natRank(rows, tipo, key, byPlan, limit) {
    const L = rows.filter((t) => t.tipo === tipo);
    const m = new Map();
    let tot = 0;
    for (const t of L) {
      const v = Math.abs(t.value);
      tot += v;
      const k = t.natK;
      if (!m.has(k)) m.set(k, { k, v: 0, n: 0, pl: {}, sub: new Map() });
      const o = m.get(k);
      o.v += v;
      o.n++;
      o.pl[t.plano] = (o.pl[t.plano] || 0) + v;
      const sk = subKey(t);
      o.sub.set(sk, (o.sub.get(sk) || 0) + v);
    }
    const list = [...m.values()].sort((a, b) => b.v - a.v);
    if (!list.length)
      return `<p class="empty-note">Sem ${tipo === "RECEITA" ? "receitas" : "despesas"} no período.</p>`;
    const max = list[0].v;
    const lim = limit || 8;
    const shown = ui.open.has(key + ":all") ? list : list.slice(0, lim);
    const col = tipo === "RECEITA" ? "var(--pos)" : "var(--neg)";
    const rowsH = shown
      .map((x, i) => {
        const { code, name } = splitCode(x.k);
        const ok = key + ":" + x.k;
        const open = ui.open.has(ok);
        const segs = byPlan
          ? PL.concat([SEM])
              .filter((p) => x.pl[p.k])
              .map(
                (p) =>
                  `<i style="width:${((x.pl[p.k] / max) * 100).toFixed(2)}%;background:${p.c};animation-delay:${i * 40}ms" title="${esc(p.l)}: ${money(x.pl[p.k])}"></i>`,
              )
              .join("")
          : `<i style="width:${((x.v / max) * 100).toFixed(2)}%;background:${col};opacity:.85;animation-delay:${i * 40}ms"></i>`;
        let sub = "";
        if (open) {
          const subs = [...x.sub.entries()].sort((a, b) => b[1] - a[1]);
          const top = subs.slice(0, 8);
          sub = `<div class="subs">${
            byPlan
              ? `<div class="t">Por plano</div>${PL.concat([SEM])
                  .filter((p) => x.pl[p.k])
                  .map(
                    (p) =>
                      `<div class="s"><span><i style="background:${p.c}"></i>${esc(p.full)}</span><b>${money(x.pl[p.k])}<small>${pct(x.pl[p.k], x.v)}</small></b></div>`,
                  )
                  .join("")}`
              : ""
          }
        <div class="t">${tipo === "RECEITA" ? "Por identificação / origem" : "Por credor"}</div>${top.map(([k, v]) => `<div class="s"><span title="${esc(k)}">${esc(k || "—")}</span><b>${money(v)}<small>${pct(v, x.v)}</small></b></div>`).join("")}${subs.length > 8 ? `<div class="s"><span>Demais (${subs.length - 8})</span><b>${money(subs.slice(8).reduce((s, e) => s + e[1], 0))}</b></div>` : ""}</div>`;
        }
        return `<div class="rk ${open ? "open" : ""}"><button type="button" data-open="${esc(ok)}" aria-expanded="${open}"><span class="lb">${CHEV}${code ? `<code>${esc(code)}</code>` : ""}<span title="${esc(name)}">${esc(name)}</span></span><span class="vl">${short(x.v)}<small>${pct(x.v, tot)}</small></span><span class="track">${segs}</span></button>${sub}</div>`;
      })
      .join("");
    const more =
      list.length > lim
        ? `<div class="more"><button class="btn sm ghost" type="button" data-open="${esc(key + ":all")}">${ui.open.has(key + ":all") ? "Mostrar menos" : `Ver as ${list.length} naturezas`}</button></div>`
        : "";
    return `<div class="rank">${rowsH}</div>${more}`;
  }

  /* aportes agrupados; byPlan = separa por plano (visão geral) */
  /* Lista dos aportes (NP e repasses), com os lançamentos ao clicar. */
  function aporteList(rows, key, byPlan) {
    const L = rows.filter((t) => t.tipo === "APORTE");
    if (!L.length) return `<p class="empty-note">Nenhum aporte no período.</p>`;
    const m = new Map();
    for (const t of L) {
      const k = (byPlan ? t.plano + "|" : "") + t.natK;
      if (!m.has(k)) m.set(k, { k, plano: t.plano, nat: t.natK, v: 0, n: 0, items: [] });
      const o = m.get(k);
      o.v += t.value;
      o.n++;
      o.items.push(t);
    }
    const list = [...m.values()].sort(
      (a, b) => PL.findIndex((p) => p.k === a.plano) - PL.findIndex((p) => p.k === b.plano) || b.v - a.v,
    );
    const max = Math.max(...list.map((x) => Math.abs(x.v))) || 1;
    return `<div class="rank">${list
      .map((x, i) => {
        const ok = key + ":" + x.k;
        const open = ui.open.has(ok);
        const p = plInfo(x.plano);
        const { code, name } = splitCode(x.nat);
        const items = x.items.slice().sort((a, b) => (a.date < b.date ? 1 : -1));
        const sub = open
          ? `<div class="subs"><div class="t">Lançamentos</div>${items
              .slice(0, 10)
              .map(
                (t) =>
                  `<div class="s"><span>${fDay(t.date)} · ${esc(subKey(t) || t.lanc || "—")}</span><b>${sgn(t.value)}</b></div>`,
              )
              .join(
                "",
              )}${items.length > 10 ? `<div class="s"><span>Demais (${items.length - 10})</span><b>${sgn(items.slice(10).reduce((s, t) => s + t.value, 0))}</b></div>` : ""}</div>`
          : "";
        return `<div class="rk ${open ? "open" : ""}"><button type="button" data-open="${esc(ok)}" aria-expanded="${open}"><span class="lb">${CHEV}${byPlan ? `<i style="width:8px;height:8px;border-radius:3px;background:${p.c};flex:none;align-self:center"></i><code>${esc(p.l)}</code>` : ""}${code ? `<code>${esc(code)}</code>` : ""}<span title="${esc(name)}">${esc(name)}</span></span><span class="vl" style="color:${x.v < 0 ? "var(--neg)" : "var(--ink)"}">${x.v < 0 ? "−" : "+"} ${short(Math.abs(x.v)).replace("−", "")}<small>${plural(x.n, "lanç.", "lanç.")}</small></span><span class="track"><i style="width:${((Math.abs(x.v) / max) * 100).toFixed(2)}%;background:${x.v < 0 ? "var(--neg)" : "var(--apo)"};opacity:.85;animation-delay:${i * 40}ms"></i></span></button>${sub}</div>`;
      })
      .join("")}</div>`;
  }

  /* ---------- telas ---------- */
  const view = () => $("#view");
  function emptyView() {
    return `<section class="glass empty"><h2>Conecte a pasta ou carregue as planilhas</h2>
    <p>O painel junta as planilhas de conciliação de todas as contas e mostra o fluxo de caixa por plano: Taxa, Financeiro, Militar e Previdenciário. Entram as linhas de Receita, Despesa e NP (aportes).</p>
    <div class="acts">${SP.ok ? '<button class="btn acc" type="button" data-go="planilhas">Conectar SharePoint</button><button class="btn" type="button" data-fold="pick">Conectar pasta</button>' : '<button class="btn acc" type="button" data-fold="pick">Conectar pasta</button>'}<button class="btn" type="button" data-act="load">Escolher arquivos</button><button class="btn ghost" type="button" data-act="sample">Ver com dados fictícios</button></div></section>`;
  }
  function headHTML(eyebrow, title, desc, from, to, dot) {
    return `<div class="head"><div><div class="eyebrow">${eyebrow}</div><h1>${dot ? `<i style="background:${dot}"></i>` : ""}${title}</h1><p>${desc}</p></div>${from ? periodCtrl(from, to) : ""}</div>`;
  }
  function sampleFlag() {
    return state.embedded
      ? `<div><span class="sample-flag">● Dados fictícios de exemplo. Carregue as planilhas para ver os números reais.</span></div>`
      : "";
  }

  /* Tela "Visão geral": cartões dos planos, gráfico, mês a mês, atenção, naturezas e aportes. */
  function renderOverview() {
    if (!ALL.length) {
      view().innerHTML =
        headHTML(
          todayLine(),
          "Fluxo de caixa",
          "Receitas e despesas das contas do Rioprevidência, separadas por plano.",
        ) + emptyView();
      return;
    }
    const { from, to, inP, pr, prev } = ctx();
    const R = rec(inP),
      A = apo(inP),
      D = des(inP),
      res = R + A - D;
    const plans = PL.concat(ALL.some((t) => t.plano === "sem") ? [SEM] : []);
    const by = (k) => inP.filter((t) => t.plano === k);
    const cards = plans
      .map((p) => {
        const L = by(p.k),
          r = rec(L),
          a = apo(L),
          d = des(L),
          x = r + a - d,
          all = ALL.filter((t) => t.plano === p.k);
        if (!all.length)
          return `<button type="button" class="plan glass empty" data-go="planilhas"><div class="top"><span class="tile" style="background:${p.c}">${ICON[p.k]}</span><div class="nm"><small>${esc(p.full)}</small><b>Nenhuma planilha</b></div></div><div class="io"><span>Defina o plano na aba Planilhas</span></div></button>`;
        return `<button type="button" class="plan glass" data-go="${p.k === "sem" ? "planilhas" : p.k}"><div class="top"><span class="tile" style="background:${p.c}">${ICON[p.k]}</span><div class="nm"><small>${esc(p.full)}</small><b title="${sgn(x)}">${big(x, true)}</b></div></div>
      <div class="io"><span class="r">Receitas <b>${short(r)}</b></span>${a ? `<span title="${money(a)}">Aportes <b style="color:var(--apo)">${a < 0 ? "−" : "+"}${short(Math.abs(a)).replace("R$ ", "")}</b></span>` : ""}<span class="d">Despesas <b>${short(d)}</b></span></div>${spark(
        all.filter((t) => t.date >= from && t.date <= to),
        from,
        to,
        p.c,
      )}</button>`;
      })
      .join("");
    const shareR = plans.map((p) => ({ l: p.l, c: p.c, v: rec(by(p.k)) }));
    const pendR = inP.filter((t) => t.tipo === "RECEITA" && t.natK === "Natureza não identificada");
    const nc = inP.filter((t) => t.conc === 0);
    const cc = inP.filter((t) => t.conc != null);
    view().innerHTML =
      headHTML(
        todayLine(),
        "Fluxo de caixa",
        `${esc(cap(periodTitle(from, to)))}: receitas de <b>${money(R)}</b>${A ? `, aportes de <b>${sgn(A)}</b>` : ""} e despesas de <b>${money(D)}</b>. Resultado de <b>${sgn(res)}</b>.`,
        from,
        to,
      ) +
      sampleFlag() +
      `<section class="plans">${cards}</section>
  <div class="grid-main">
    <section class="glass card"><div class="card-h"><div><h2>Receitas × despesas</h2><div class="big"><b>${sgn(res)}</b>${deltaChip(res, pr ? resOf(prev) : 0, true, pr && pr.lab)}</div><div class="sub">Resultado no período (receitas + aportes − despesas) · ${fDay(from)} a ${fDay(to)}</div></div>
      <div class="legend"><span><i style="background:var(--accent)"></i>Receitas</span>${A ? '<span><i style="background:var(--apo)"></i>Aportes</span>' : ""}<span><i class="dash" style="border-color:var(--neg)"></i>Despesas</span></div></div>
      <div class="chart" id="chart"></div>
      ${shareBar(shareR, R)}
    </section>
    <div class="col">
      <section class="glass card"><div class="card-h"><div><h2>Resultado mês a mês</h2><div class="sub">Receitas + aportes − despesas, todos os planos</div></div></div>${monthBars(ALL, to)}</section>
      <section class="glass card"><div class="card-h"><div><h2>Atenção</h2><div class="sub">No período escolhido</div></div></div><div class="att">
        <div class="it"><span class="ic">${ICON.sem}</span><div><span class="big2">${INT.format(pendR.length)}</span><b>${pendR.length === 1 ? "receita sem natureza" : "receitas sem natureza"}</b><span>${money(rec(pendR))} ainda como "Não encontrado"</span></div></div>
        <div class="it"><span class="ic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg></span><div><span class="big2">${cc.length ? pct(cc.length - nc.length, cc.length) : "—"}</span><b>conciliado</b><span>${cc.length ? `${plural(nc.length, "lançamento pendente", "lançamentos pendentes")}` : "a planilha não informa a conciliação"}</span></div></div>
      </div></section>
    </div>
  </div>
  <div class="grid-2">
    <section class="glass card"><div class="card-h"><div><h2>Receitas por natureza</h2><div class="sub">${money(R)} · cores por plano · clique para ver as subdivisões</div></div></div>${natRank(inP, "RECEITA", "gR", true)}</section>
    <section class="glass card"><div class="card-h"><div><h2>Despesas por natureza</h2><div class="sub">${money(D)} · cores por plano · clique para ver as subdivisões</div></div></div>${natRank(inP, "DESPESA", "gD", true)}</section>
  </div>
  <section class="glass card"><div class="card-h"><div><h2>Aportes</h2><div class="sub">NP e repasses do Financeiro para o Militar · líquido ${sgn(A)} · clique para ver os lançamentos</div></div></div>${aporteList(inP, "gA", true)}</section>`;
    drawChart($("#chart"), inP, from, to, "var(--accent)");
  }
  /* Tela de um plano: indicadores, gráfico, naturezas, aportes e lançamentos. */
  function renderPlano(k) {
    const p = plInfo(k);
    const own = ALL.filter((t) => t.plano === k);
    const nSheets = ORDER.filter((n) => (planoOf(state.accounts[n]) || "sem") === k).length;
    if (!own.length) {
      view().innerHTML =
        headHTML(
          p.full,
          p.l,
          ALL.length
            ? `Nenhuma planilha está marcada como ${esc(p.full)}. Defina o plano de cada planilha na aba Planilhas.`
            : "Ainda não há planilhas carregadas.",
          "",
          "",
          p.c,
        ) +
        (ALL.length
          ? `<section class="glass empty"><h2>Sem planilhas neste plano</h2><p>O plano é reconhecido pelo nome do arquivo ou da pasta (Taxa ou Adm, Fin, Mil, Prev) e pela conta contábil dentro da planilha. Se não reconheceu, escolha na aba Planilhas.</p><div class="acts"><button class="btn acc" type="button" data-go="planilhas">Abrir Planilhas</button></div></section>`
          : emptyView());
      return;
    }
    const { from, to, pr } = ctx();
    const inP = own.filter((t) => t.date >= from && t.date <= to),
      prev = pr ? own.filter((t) => t.date >= pr.pf && t.date <= pr.pt) : [];
    const R = rec(inP),
      A = apo(inP),
      D = des(inP),
      res = R + A - D;
    const aIn = apoIn(inP),
      aOut = apoOut(inP);
    const nats = (() => {
      const m = new Map();
      for (const t of inP.filter((t) => t.tipo === "RECEITA")) {
        m.set(t.natK, (m.get(t.natK) || 0) + t.value);
      }
      return [...m.entries()].sort((a, b) => b[1] - a[1]);
    })();
    const palette = [p.c, "#CFE3F5", "#94B3D1", "#62809E", "#3E5872"];
    const share = nats
      .slice(0, 4)
      .map(([k2, v], i) => ({ l: splitCode(k2).name, c: palette[i], v }))
      .concat(
        nats.length > 4
          ? [{ l: "Demais naturezas", c: palette[4], v: nats.slice(4).reduce((s, e) => s + e[1], 0) }]
          : [],
      );
    const pendR = inP.filter((t) => t.tipo === "RECEITA" && t.natK === "Natureza não identificada");
    const cc = inP.filter((t) => t.conc != null),
      ok = cc.filter((t) => t.conc === 1).length;
    view().innerHTML =
      headHTML(
        `${esc(p.full)} · ${plural(nSheets, "planilha", "planilhas")}`,
        p.l,
        `${esc(cap(periodTitle(from, to)))}: receitas de <b>${money(R)}</b>${A ? `, aportes de <b>${sgn(A)}</b>` : ""} e despesas de <b>${money(D)}</b>.`,
        from,
        to,
        p.c,
      ) +
      sampleFlag() +
      `<section class="kpis">
    <div class="glass kpi"><div class="lab"><i style="background:var(--pos)"></i>Receitas</div><div class="val" title="${money(R)}">${big(R)}</div><div class="foot">${deltaChip(R, pr ? rec(prev) : 0, true, pr && pr.lab)}</div></div>
    <div class="glass kpi"><div class="lab"><i style="background:var(--apo)"></i>Aportes</div><div class="val" title="${sgn(A)}">${A ? big(A, true) : money(0)}</div><div class="foot">${aIn ? `<span>entraram ${short(aIn)}</span>` : ""}${aOut ? `<span>${k === "financeiro" ? "repassados ao Militar" : "saíram"} ${short(aOut)}</span>` : ""}${!aIn && !aOut ? (k === "militar" || k === "financeiro" ? '<button type="button" class="btn sm ghost" data-act="toap" style="padding:2px 8px">Marcar as transferências</button>' : "<span>nenhum no período</span>") : ""}</div></div>
    <div class="glass kpi"><div class="lab"><i style="background:var(--neg)"></i>Despesas</div><div class="val" title="${money(D)}">${big(D)}</div><div class="foot">${deltaChip(D, pr ? des(prev) : 0, false, pr && pr.lab)}</div></div>
    <div class="glass kpi"><div class="lab"><i style="background:${p.c}"></i>Resultado</div><div class="val" style="color:${res < 0 ? "var(--neg)" : "var(--ink)"}" title="${sgn(res)}">${big(res, true)}</div><div class="foot">receitas + aportes − despesas</div></div>
    <div class="glass kpi ${pendR.length ? "warn" : ""}"><div class="lab">${pendR.length ? "Receitas sem natureza" : "Conciliado"}</div><div class="val">${pendR.length ? INT.format(pendR.length) : cc.length ? pct(ok, cc.length) : "—"}</div><div class="foot">${pendR.length ? money(rec(pendR)) : cc.length ? `${INT.format(ok)} de ${INT.format(cc.length)}` : ""}</div></div>
  </section>
  <div class="grid-main">
    <section class="glass card"><div class="card-h"><div><h2>Receitas × despesas</h2><div class="sub">${fDay(from)} a ${fDay(to)}</div></div><div class="legend"><span><i style="background:${p.c}"></i>Receitas</span>${aIn ? '<span><i style="background:var(--apo)"></i>Aportes</span>' : ""}<span><i class="dash" style="border-color:var(--neg)"></i>Despesas</span></div></div>
      <div class="chart" id="chart"></div>${shareBar(share, R)}</section>
    <section class="glass card"><div class="card-h"><div><h2>Resultado mês a mês</h2><div class="sub">${esc(p.full)}</div></div></div>${monthBars(own, to)}</section>
  </div>
  <div class="grid-2">
    <section class="glass card"><div class="card-h"><div><h2>Receitas por natureza</h2><div class="sub">${money(R)} · clique para ver as subdivisões</div></div></div>${natRank(inP, "RECEITA", k + "R", false, 10)}</section>
    <section class="glass card"><div class="card-h"><div><h2>Despesas por natureza</h2><div class="sub">${money(D)} · clique para ver as subdivisões</div></div></div>${natRank(inP, "DESPESA", k + "D", false, 10)}</section>
  </div>
  ${
    inP.some((t) => t.tipo === "APORTE") || k === "militar" || k === "financeiro"
      ? `<section class="glass card" id="apCard"><div class="card-h"><div><h2>Aportes</h2><div class="sub">${k === "financeiro" ? "NP recebidas e repasses para o Militar" : k === "militar" ? "NP e aportes recebidos do Financeiro" : "Linhas NP"} · líquido ${sgn(A)}</div></div></div>${aporteList(inP, k + "A", false)}
    ${k === "militar" || k === "financeiro" ? `<details class="pick-ap" ${own.some((t) => t.tipo === "APORTE") ? "" : "open"} style="margin-top:14px;border-top:1px solid var(--line);padding-top:12px"><summary style="cursor:pointer;font-weight:600;font-size:13.5px">${own.some((t) => t.tipo === "APORTE") ? "Mudar quais transferências são aporte" : "Escolha quais transferências são o aporte do Financeiro para o Militar"}</summary><p style="color:var(--muted);font-size:12.5px;margin:6px 0 12px">Marque e o painel atualiza na hora. A escolha fica guardada.</p>${destBody(k)}</details>` : ""}</section>`
      : ""
  }
  <section class="glass card"><div class="card-h"><div><h2>Lançamentos</h2><div class="sub" id="txSum"></div></div><label class="search"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg><input type="search" id="q" placeholder="Buscar natureza, ente, credor, aporte, OB…" aria-label="Buscar lançamentos" value="${esc(ui.q)}"></label></div>
    <div class="tbl-wrap"><table class="tx"><thead><tr><th>Data</th><th>Tipo</th><th>Natureza</th><th>Identificação / credor</th><th>Doc.</th><th class="r">Valor</th></tr></thead><tbody id="txBody"></tbody></table></div><div class="more" id="txMore"></div></section>`;
    drawChart($("#chart"), inP, from, to, p.c);
    renderTx(inP);
  }
  let txRows = [];
  /* Tabela de lançamentos da tela do plano (com busca). */
  function renderTx(rows) {
    if (rows) txRows = rows;
    const q = norm(ui.q);
    const L = txRows
      .filter((t) => !q || t.s.includes(q))
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : b.id - a.id));
    const shown = L.slice(0, ui.limit);
    $("#txBody").innerHTML =
      shown
        .map((t) => {
          const { code, name } = splitCode(t.natK);
          const ap = t.tipo === "APORTE";
          const who =
            t.tipo === "RECEITA" ? t.ident || t.lanc : ap ? subKey(t) : splitCode(t.credor).name || t.lanc;
          const sub = t.tipo === "RECEITA" ? t.idk || t.lanc : ap ? t.lanc : t.item || t.lanc;
          const tag = ap
            ? t.natK === APO_MIL_FIN
              ? "Aporte do Financeiro"
              : t.natK === APO_FIN_MIL
                ? "Repasse ao Militar"
                : t.value >= 0
                  ? "Aporte (NP)"
                  : "Saída de aporte"
            : t.tipo === "RECEITA"
              ? "Receita"
              : "Despesa";
          return `<tr><td class="d">${fDay(t.date)}</td><td><span class="tag ${ap ? "A" : t.tipo[0]}">${tag}</span></td><td><span class="m">${esc(name)}</span>${code ? `<span class="s">${esc(code)}</span>` : ""}</td><td><span class="m">${esc(who || "—")}</span>${sub && sub !== who ? `<span class="s">${esc(sub)}</span>` : ""}</td><td class="d">${esc(t.doc || "—")}</td><td class="r v ${t.value > 0 ? "p" : "n"}">${t.value > 0 ? "+" : "−"} ${money(Math.abs(t.value))}</td></tr>`;
        })
        .join("") ||
      `<tr><td colspan="6" style="color:var(--muted);padding:18px 10px">Nenhum lançamento com essa busca.</td></tr>`;
    $("#txSum").textContent = `${plural(L.length, "lançamento", "lançamentos")} · resultado ${sgn(resOf(L))}`;
    const rest = L.length - shown.length;
    $("#txMore").innerHTML =
      rest > 0
        ? `<button class="btn sm" type="button" data-act="more">Mostrar mais ${INT.format(Math.min(rest, 100))} de ${INT.format(rest)}</button>`
        : "";
  }

  /* ---------- relatório para imprimir ---------- */
  /* Tabela de naturezas do relatório impresso. */
  function natTable(L, tipo, detail) {
    const sign = tipo === "RECEITA" ? 1 : -1;
    const m = new Map();
    let tot = 0;
    for (const t of L.filter((t) => t.tipo === tipo)) {
      const v = sign * t.value;
      tot += v;
      if (!m.has(t.natK)) m.set(t.natK, { v: 0, n: 0, sub: new Map() });
      const o = m.get(t.natK);
      o.v += v;
      o.n++;
      const sk = subKey(t);
      o.sub.set(sk, (o.sub.get(sk) || 0) + v);
    }
    if (!m.size) return `<p class="note">Sem ${tipo === "RECEITA" ? "receitas" : "despesas"} no período.</p>`;
    const rows = [...m.entries()]
      .sort((a, b) => b[1].v - a[1].v)
      .map(([k, o]) => {
        const { code, name } = splitCode(k);
        let h = `<tr><td class="c">${esc(code)}</td><td>${esc(name)}</td><td class="r">${INT.format(o.n)}</td><td class="r">${money(o.v)}</td><td class="r">${pct(o.v, tot)}</td></tr>`;
        if (detail) {
          const subs = [...o.sub.entries()].sort((a, b) => b[1] - a[1]);
          h +=
            subs
              .slice(0, 6)
              .map(
                ([s, v]) =>
                  `<tr class="sub"><td></td><td>${esc(s || "—")}</td><td></td><td class="r">${money(v)}</td><td class="r">${pct(v, o.v)}</td></tr>`,
              )
              .join("") +
            (subs.length > 6
              ? `<tr class="sub"><td></td><td>Demais (${subs.length - 6})</td><td></td><td class="r">${money(subs.slice(6).reduce((s, e) => s + e[1], 0))}</td><td></td></tr>`
              : "");
        }
        return h;
      })
      .join("");
    return `<div class="rt-wrap"><table><thead><tr><th>Código</th><th>Natureza ${tipo === "RECEITA" ? "da receita" : "da despesa"}</th><th class="r">Qtde</th><th class="r">Valor</th><th class="r">%</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td></td><td>Total de ${tipo === "RECEITA" ? "receitas" : "despesas"}</td><td class="r">${INT.format(L.filter((t) => t.tipo === tipo).length)}</td><td class="r">${money(tot)}</td><td class="r">100%</td></tr></tfoot></table></div>`;
  }
  /* Tabela de aportes do relatório impresso. */
  function aporteTable(L, detail) {
    const A = L.filter((t) => t.tipo === "APORTE");
    if (!A.length) return '<p class="note">Sem aportes no período.</p>';
    const m = new Map();
    for (const t of A) {
      if (!m.has(t.natK)) m.set(t.natK, { v: 0, n: 0, items: [] });
      const o = m.get(t.natK);
      o.v += t.value;
      o.n++;
      o.items.push(t);
    }
    const rows = [...m.entries()]
      .sort((a, b) => b[1].v - a[1].v)
      .map(([k, o]) => {
        const { code, name } = splitCode(k);
        let h = `<tr><td class="c">${esc(code)}</td><td>${esc(name)}</td><td class="r">${INT.format(o.n)}</td><td class="r ${o.v < 0 ? "neg" : ""}">${sgn(o.v)}</td></tr>`;
        if (detail)
          h += o.items
            .slice()
            .sort((a, b) => (a.date < b.date ? -1 : 1))
            .slice(0, 12)
            .map(
              (t) =>
                `<tr class="sub"><td></td><td>${fDay(t.date)} · ${esc(subKey(t) || t.lanc || "—")}</td><td></td><td class="r">${sgn(t.value)}</td></tr>`,
            )
            .join("");
        return h;
      })
      .join("");
    return `<div class="rt-wrap"><table><thead><tr><th>Código</th><th>Aporte</th><th class="r">Qtde</th><th class="r">Valor</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td></td><td>Total de aportes (líquido)</td><td class="r">${INT.format(A.length)}</td><td class="r">${sgn(apo(L))}</td></tr></tfoot></table></div>`;
  }
  function months(from, to) {
    const out = [];
    let k = from.slice(0, 8) + "01";
    while (k <= to) {
      out.push(k.slice(0, 7));
      const d = toD(k);
      d.setMonth(d.getMonth() + 1);
      k = fromD(d);
    }
    return out;
  }
  /* Tela "Relatórios": documento A4 do fluxo de caixa por plano, pronto para imprimir. */
  function renderReport() {
    if (!ALL.length) {
      view().innerHTML =
        headHTML(
          "Para imprimir",
          "Relatórios",
          "Carregue as planilhas para gerar o relatório de fluxo de caixa por plano.",
        ) + emptyView();
      return;
    }
    const { from, to, inP } = ctx();
    const plans = PL.concat(inP.some((t) => t.plano === "sem") ? [SEM] : []);
    const sel = ui.rplan === "todos" ? plans : plans.filter((p) => p.k === ui.rplan);
    const now = new Date();
    const files = [...new Set(ORDER.map((n) => state.accounts[n].file).filter(Boolean))].length;
    const scopeBtns = [["todos", "Todos os planos"]].concat(PL.map((p) => [p.k, p.l]));
    let doc = `<header><div class="org"><img src="${LOGO}" alt="">Rioprevidência · Tesouraria</div><h2>Fluxo de Caixa ${ui.rplan === "todos" ? "por Plano" : "· " + esc(plInfo(ui.rplan).full)}</h2>
    <div class="dmeta"><div><b>Período</b>${esc(cap(periodTitle(from, to)))}</div><div><b>Planos</b>${esc(sel.map((p) => p.l).join(", "))}</div><div><b>Emitido em</b>${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} às ${pad(now.getHours())}h${pad(now.getMinutes())}</div><div><b>Base</b>Receitas, despesas e aportes (NP e transferências do Financeiro para o Militar)</div></div></header>`;
    if (ui.rplan === "todos") {
      const tR = rec(inP),
        tA = apo(inP),
        tD = des(inP),
        tX = tR + tA - tD;
      doc += `<section><h3>Resumo por plano</h3><p>Entre ${fDayL(from)} e ${fDayL(to)}, os planos tiveram ${money(tR)} de receitas${tA ? `, ${sgn(tA)} de aportes` : ""} e ${money(tD)} de despesas, com resultado de <b class="${tX < 0 ? "neg" : "pos"}">${sgn(tX)}</b>. Os repasses do Financeiro para o Militar aparecem como saída no Financeiro e entrada no Militar, e se anulam no total.</p>
      <div class="rt-wrap"><table><thead><tr><th>Plano</th><th class="r">Receitas</th><th class="r">Aportes</th><th class="r">Despesas</th><th class="r">Resultado</th><th class="r">% das receitas</th></tr></thead><tbody>${plans
        .map((p) => {
          const L = inP.filter((t) => t.plano === p.k),
            r = rec(L),
            a = apo(L),
            d = des(L),
            x = r + a - d;
          return `<tr><td>${esc(p.full)}</td><td class="r">${money(r)}</td><td class="r ${a < 0 ? "neg" : ""}">${a ? sgn(a) : "—"}</td><td class="r">${money(d)}</td><td class="r ${x < 0 ? "neg" : "pos"}">${sgn(x)}</td><td class="r">${pct(r, tR)}</td></tr>`;
        })
        .join(
          "",
        )}</tbody><tfoot><tr><td>Total</td><td class="r">${money(tR)}</td><td class="r">${tA ? sgn(tA) : "—"}</td><td class="r">${money(tD)}</td><td class="r">${sgn(tX)}</td><td class="r">100%</td></tr></tfoot></table></div></section>`;
      const ms = months(from, to);
      if (ms.length > 1)
        doc += `<section><h3>Resultado mês a mês</h3><div class="rt-wrap"><table><thead><tr><th>Mês</th>${plans.map((p) => `<th class="r">${esc(p.l)}</th>`).join("")}<th class="r">Total</th></tr></thead><tbody>${ms
          .map((m) => {
            const L = inP.filter((t) => t.date.startsWith(m));
            return `<tr><td>${MESES_L[+m.slice(5) - 1]}/${m.slice(0, 4)}</td>${plans
              .map((p) => {
                const x = L.filter((t) => t.plano === p.k);
                const v = resOf(x);
                return `<td class="r ${v < 0 ? "neg" : ""}">${x.length ? sgn(v) : "—"}</td>`;
              })
              .join("")}<td class="r">${sgn(resOf(L))}</td></tr>`;
          })
          .join("")}</tbody></table></div></section>`;
    }
    for (const p of sel) {
      const L = inP.filter((t) => t.plano === p.k);
      const x = resOf(L);
      const hasA = L.some((t) => t.tipo === "APORTE");
      doc += `<section class="plano-sec"><h3><i style="background:${p.c}"></i>${esc(p.full)}</h3>${L.length ? `<h4>Receitas por natureza</h4>${natTable(L, "RECEITA", ui.rsub)}${hasA ? `<h4>Aportes</h4>${aporteTable(L, ui.rsub)}` : ""}<h4>Despesas por natureza</h4>${natTable(L, "DESPESA", ui.rsub)}<div class="res"><span>Resultado do ${esc(p.l === "Taxa" ? "plano Taxa" : p.full.replace("Plano ", "plano "))} (receitas${hasA ? " + aportes" : ""} − despesas)</span><span class="${x < 0 ? "neg" : "pos"}">${sgn(x)}</span></div>` : '<p class="note">Sem lançamentos de receita, despesa ou aporte no período.</p>'}</section>`;
    }
    doc += `<footer><span>${plural(files, "planilha de conciliação", "planilhas de conciliação")} · receitas, despesas e aportes (NP e transferências do Financeiro para o Militar); as demais transferências e o FI ficam de fora</span><span>Valores em reais</span></footer>`;
    view().innerHTML =
      headHTML(
        "Para imprimir",
        "Relatórios",
        "Fluxo de caixa por plano, com receitas e despesas por natureza e os aportes à parte. Escolha o período e o plano, depois imprima.",
        from,
        to,
      ) +
      sampleFlag() +
      `<div id="docBox" style="display:grid;gap:14px"><div class="rbar glass card" style="padding:12px 14px"><div class="ctrl"><div class="seg" role="group" aria-label="Plano do relatório">${scopeBtns.map(([k, l]) => `<button type="button" data-rplan="${k}" aria-pressed="${ui.rplan === k}">${esc(l)}</button>`).join("")}</div><label><input type="checkbox" id="rsub" ${ui.rsub ? "checked" : ""}> Detalhar cada natureza (identificação e credor)</label></div>
    <div class="ctrl">${FRAMED ? '<span style="font-size:12.5px;color:var(--muted)">Para imprimir, abra o arquivo no Edge ou no Chrome.</span>' : '<button class="btn acc" type="button" data-act="print"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V3h12v6M6 18H4a1 1 0 0 1-1-1v-6a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v6a1 1 0 0 1-1 1h-2M6 14h12v7H6z"/></svg>Imprimir relatório</button>'}<button class="btn" type="button" data-act="copy">Copiar texto</button></div></div>
    <article class="paper" id="doc">${doc}</article></div>`;
  }

  /* ---------- planilhas ---------- */
  /* Tela "Planilhas": fontes (SharePoint, pasta, upload), plano de cada planilha e aportes. */
  function renderSheetsView() {
    const miss = ORDER.filter((n) => !planoOf(state.accounts[n])).length;
    const rowsH = ORDER.map((nm) => {
      const a = state.accounts[nm];
      const p = planoOf(a);
      const set = (state.planoSet || {})[a.srcName || a.name];
      const L = a.txs.filter((t) => t.tipo === "RECEITA" || t.tipo === "DESPESA" || t.tipo === "NP");
      const d0 = L.length ? L[0].date : null,
        d1 = L.length ? L[L.length - 1].date : null;
      const nr = L.filter((t) => t.tipo === "RECEITA").length,
        na = L.filter((t) => t.tipo === "NP").length,
        nd = L.length - nr - na;
      const byName = planoFromName((a.paths || []).join(" ") + " " + (a.file || "") + " " + a.name);
      const how = set
        ? "escolhido aqui"
        : p && byName === p
          ? a.src === "pasta"
            ? "reconhecido pela pasta ou nome"
            : "reconhecido pelo nome"
          : p
            ? "reconhecido na planilha"
            : "";
      return `<tr class="${p ? "" : "miss"}" data-a="${esc(nm)}"><td><b>${esc(nm)}</b><span class="s">${esc(a.file || "")}${a.src === "pasta" ? " · da pasta" : ""}</span></td>
      <td class="mono" style="white-space:nowrap;font-size:12px;color:var(--muted)">${d0 ? `${fDay(d0)} a ${fDay(d1)}` : "—"}</td>
      <td style="white-space:nowrap">${INT.format(nr)} receitas<span class="s">${INT.format(nd)} despesas${na ? ` · ${INT.format(na)} NP` : ""}${(a.pv || 1) < PARSER_V ? ' · <span style="color:var(--warn)">carregue de novo para incluir as NP</span>' : ""}</span></td>
      <td><select data-plano aria-label="Plano de ${esc(nm)}"><option value="">Escolher plano…</option>${PL.map((x) => `<option value="${x.k}" ${p === x.k ? "selected" : ""}>${esc(x.full)}</option>`).join("")}</select>${how ? `<span class="s srcs">${how}</span>` : ""}</td>
      <td style="text-align:right">${a.src === "pasta" ? "" : `<button class="btn sm ghost" type="button" data-del>Remover</button>`}</td></tr>`;
    }).join("");
    view().innerHTML =
      headHTML(
        "Fonte dos dados",
        "Planilhas",
        `${plural(ORDER.length, "planilha carregada", "planilhas carregadas")}${miss ? ` · <b style="color:var(--warn)">${plural(miss, "sem plano", "sem plano")}</b>` : ""}. O painel não separa por conta corrente: cada planilha entra no plano escolhido aqui.`,
      ) +
      spCardHTML() +
      `<section class="glass folder" id="folder"><span class="fic" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg></span>
    <div class="fbody"><div class="ftop"><b>Pasta no computador (OneDrive)</b><span id="fstat"></span></div><div class="finfo" id="finfo"></div></div><div class="acts" id="facts"></div><div class="ferr" id="ferr"></div></section>
  <section class="glass drop-zone" data-act="load" tabindex="0" role="button"><span class="fic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4M6 10l6-6 6 6"/><path d="M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3"/></svg></span><div><b>Adicionar planilhas</b><span>Clique ou solte aqui. Aceita a planilha de conciliação (.xlsm/.xlsx, aba C.C) e extratos em .csv.</span></div></section>
  ${
    ORDER.length
      ? `<section class="glass card"><div class="card-h"><div><h2>Planilhas e planos</h2><div class="sub">Confira o plano de cada planilha. As que estão em amarelo ainda não têm plano.</div></div>${state.embedded ? "" : `<button class="btn sm ghost" type="button" data-act="clear">Limpar tudo</button>`}</div>
    <div class="tbl-wrap"><table class="sheets"><thead><tr><th>Planilha</th><th>Período</th><th>Lançamentos</th><th>Plano</th><th></th></tr></thead><tbody>${rowsH}</tbody></table></div></section>`
      : ""
  }
  ${destCard()}
  <section class="glass card help"><b>Como o plano é reconhecido</b><ul>
    <li>Pelo nome do arquivo ou da pasta: <b>Taxa</b> ou <b>Adm</b>, <b>Fin</b> ou Financeiro, <b>Mil</b> ou Militar, <b>Prev</b> ou Previdenciário.</li>
    <li>Pelo conteúdo da planilha: "Plano Financeiro" nas observações, conta contábil "Fundo em Repartição" (Financeiro) ou "Fundo em Capitalização" (Previdenciário), "Taxa de Administração", "Plano Militar".</li>
    <li>Se não reconhecer, escolha aqui. A escolha fica guardada e vale para as próximas cargas da mesma planilha.</li>
    <li>Entram as linhas de <b>Receita</b> e <b>Despesa</b>, e como <b>Aporte</b> (separado da receita) as linhas <b>NP</b> e as transferências do Financeiro para as contas do Militar marcadas acima. As demais transferências e as aplicações ou resgates (FI) ficam de fora.</li></ul></section>`;
    renderSP();
    renderFolder();
  }
  /* Agrupa as transferências por conta de origem/destino para a escolha dos aportes. */
  function transfGroups(plano, sign) {
    const m = new Map();
    for (const nm of ORDER) {
      const a = state.accounts[nm];
      if ((planoOf(a) || "") !== plano) continue;
      for (const t of a.txs) {
        if (t.tipo !== "TRANSF" || !t.credor || !(sign < 0 ? t.value < 0 : t.value > 0)) continue;
        const k = destKey(t.credor);
        if (!m.has(k)) m.set(k, { k, label: t.credor, n: 0, v: 0, last: "", accts: new Set() });
        const o = m.get(k);
        o.n++;
        o.v += Math.abs(t.value);
        if (t.date > o.last) o.last = t.date;
        o.accts.add(nm);
      }
    }
    return [...m.values()].sort((a, b) => b.v - a.v);
  }
  function destTable(list, attr, isOn, note, head) {
    if (!list.length) return "";
    return `<div class="tbl-wrap"><table class="sheets"><thead><tr><th>É aporte?</th><th>${head}</th><th>Planilha</th><th>Transferências</th><th style="text-align:right">Total</th></tr></thead><tbody>${list
      .map((o) => {
        const st = isOn(o);
        return `<tr><td><label style="display:inline-flex;gap:8px;align-items:center;cursor:pointer;white-space:nowrap"><input type="checkbox" ${attr}="${esc(o.k)}" ${st.on ? "checked" : ""} style="accent-color:var(--accent);width:16px;height:16px"> ${st.on ? "Sim" : "Não"}</label>${st.note ? `<span class="s srcs">${st.note}</span>` : ""}</td><td class="mono">${esc(o.label)}</td><td><span class="s" style="color:var(--muted);font-size:12px">${esc([...o.accts].join(", "))}</span></td><td>${INT.format(o.n)}<span class="s">última em ${fDay(o.last)}</span></td><td style="text-align:right;font-variant-numeric:tabular-nums">${money(o.v)}</td></tr>`;
      })
      .join("")}</tbody></table></div>`;
  }
  /* which: 'all' (aba Planilhas), 'militar' ou 'financeiro' (telas dos planos) */
  /* Quadro de escolha: quais transferências são o aporte do Financeiro para o Militar. */
  function destBody(which) {
    const nums = militarNums();
    const outF = transfGroups("financeiro", -1),
      inM = transfGroups("militar", 1);
    const chips =
      nums.auto
        .map((n) => `<span class="pill ok" title="Do nome de uma planilha do Militar">${esc(n)}</span>`)
        .join("") +
      nums.man
        .map(
          (n) =>
            `<span class="pill ok">${esc(n)} <button type="button" data-unnum="${esc(n)}" aria-label="Tirar a conta ${esc(n)}" style="background:none;border:0;color:inherit;cursor:pointer;padding:0 0 0 2px">✕</button></span>`,
        )
        .join("");
    const tIn = destTable(
      inM,
      "data-orig",
      (o) => {
        const on = !!(state.aporteOrig || {})[o.k];
        return { on, note: on ? "marcado aqui" : "" };
      },
      "",
      "Conta de origem",
    );
    const tOut = destTable(
      outF,
      "data-dest",
      (o) => {
        const ex = (state.aporteDest || {})[o.k];
        const sug = destSuggested(o.label, nums.all);
        const on = ex != null ? ex : sug;
        return {
          on,
          note: ex == null && sug ? "sugerido pela conta do Militar" : ex != null ? "marcado aqui" : "",
        };
      },
      "",
      "Conta de destino",
    );
    let h = "";
    if (which !== "financeiro")
      h += `<h3 style="font-size:13.5px;font-weight:600;margin:4px 0 8px">Transferências recebidas nas planilhas do Militar</h3>${tIn || '<p class="empty-note">Nenhuma transferência de entrada nas planilhas do Militar.</p>'}`;
    if (which !== "militar")
      h += `<h3 style="font-size:13.5px;font-weight:600;margin:16px 0 8px">Transferências que saem das planilhas do Financeiro</h3>
    <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px;font-size:12.5px;color:var(--muted)">Contas do Militar: ${chips || "<span>nenhuma ainda</span>"}
      <label style="display:inline-flex;gap:6px;align-items:center;margin-left:6px">Adicionar conta <input id="apNum" inputmode="numeric" placeholder="ex.: 292061" style="width:130px;background:rgba(0,0,0,.25);border:1px solid var(--line);border-radius:999px;padding:5px 10px;font-size:12.5px"></label><button class="btn sm" type="button" data-act="addnum">Adicionar</button></div>
    ${tOut || '<p class="empty-note">Nenhuma transferência de saída nas planilhas do Financeiro.</p>'}`;
    return h;
  }
  function destCard() {
    return `<section class="glass card"><div class="card-h"><div><h2>Aportes do Financeiro para o Militar</h2><div class="sub">Das transferências, só entram as do aporte do Financeiro para o Militar. Marque nas planilhas do Militar a entrada, ou nas do Financeiro a saída: o painel mostra a entrada no Militar e o repasse no Financeiro sem contar duas vezes.</div></div></div>${destBody("all")}</section>`;
  }
  function renderBadge() {
    const miss = ORDER.filter((n) => !planoOf(state.accounts[n])).length;
    const b = $("#navBadge");
    b.hidden = !miss;
    b.textContent = miss;
    b.title = miss ? plural(miss, "planilha sem plano", "planilhas sem plano") : "";
  }

  /* ---------- navegação ---------- */
  const VIEWS = ["geral", "taxa", "financeiro", "militar", "previdenciario", "relatorios", "planilhas"];
  let pillAnim = null,
    firstPill = true;
  /* Animação da pílula da barra de navegação (a borda da frente chega antes da de trás). */
  function movePill() {
    const nav = $("#nav"),
      a = nav.querySelector(`a[data-v="${ui.view}"]`),
      pill = $("#navPill");
    if (!a) return;
    nav.querySelectorAll("a").forEach((x) => x.toggleAttribute("aria-current", false));
    a.setAttribute("aria-current", "page");
    const toL = a.offsetLeft,
      toW = a.offsetWidth;
    const fromL = parseFloat(pill.style.left) || 0,
      fromW = parseFloat(pill.style.width) || 0;
    if (a.scrollIntoView && nav.scrollWidth > nav.clientWidth)
      nav.scrollTo({ left: toL - nav.clientWidth / 2 + toW / 2, behavior: firstPill ? "auto" : "smooth" });
    const reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (firstPill || reduce || !fromW) {
      pill.style.left = toL + "px";
      pill.style.width = toW + "px";
      firstPill = false;
      return;
    }
    // a borda da frente chega primeiro e a de trás alcança depois: a pílula estica e encolhe
    cancelAnimationFrame(pillAnim);
    const right = toL > fromL;
    const t0 = performance.now(),
      dur = 460;
    const ease = (t) => 1 - Math.pow(1 - t, 3);
    const fL = fromL,
      fR = fromL + fromW,
      tL = toL,
      tR = toL + toW;
    const step = (now) => {
      const t = Math.min(1, (now - t0) / dur);
      const lead = ease(Math.min(1, t / 0.62)),
        trail = ease(t);
      const L = right ? fL + (tL - fL) * trail : fL + (tL - fL) * lead;
      const R = right ? fR + (tR - fR) * lead : fR + (tR - fR) * trail;
      pill.style.left = L + "px";
      pill.style.width = Math.max(10, R - L) + "px";
      if (t < 1) pillAnim = requestAnimationFrame(step);
    };
    pillAnim = requestAnimationFrame(step);
  }
  /* Troca de tela pelo endereço (#geral, #financeiro, #relatorios…). */
  function route(animate) {
    const h = location.hash.slice(1);
    ui.view = VIEWS.includes(h) ? h : "geral";
    ui.limit = 50;
    render(animate);
    movePill();
  }
  /* Desenha a tela atual (com a transição de desfoque). */
  function render(animate) {
    const v = ui.view;
    const det = document.querySelector("details.pick-ap");
    const wasOpen = det ? det.open : null;
    if (v === "geral") renderOverview();
    else if (v === "relatorios") renderReport();
    else if (v === "planilhas") renderSheetsView();
    else renderPlano(v);
    if (wasOpen != null) {
      const d = document.querySelector("details.pick-ap");
      if (d) d.open = wasOpen;
    }
    if (animate && !matchMedia("(prefers-reduced-motion: reduce)").matches) {
      view().animate(
        [
          { opacity: 0, filter: "blur(14px)", transform: "translateY(8px) scale(.995)" },
          { opacity: 1, filter: "blur(0)", transform: "none" },
        ],
        { duration: 480, easing: "cubic-bezier(.2,.8,.2,1)" },
      );
    }
    $("#btnSync").hidden = !(PASTA.handle || state.folderName || SP.cfg.active);
    $("#btnSync").disabled = PASTA.lendo || SP.lendo;
    savePrefs();
  }
  const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);
  function todayLine() {
    const d = new Date();
    return `${cap(DIAS_L[d.getDay()])}, ${d.getDate()} de ${MESES_L[d.getMonth()]}`;
  }

  /* ---------- leitura de arquivos ---------- */
  function decode(buf) {
    let t;
    try {
      t = new TextDecoder("utf-8", { fatal: true }).decode(buf);
    } catch (e) {
      t = new TextDecoder("windows-1252").decode(buf);
    }
    return t.replace(/^﻿/, "");
  }
  /* Lê CSV (separador ; , ou tab), respeitando aspas. */
  function parseCSV(text) {
    const head = text.split(/\r?\n/).slice(0, 8).join("\n");
    const d = head.includes(";") ? ";" : head.includes("\t") ? "\t" : ",";
    const rows = [];
    let row = [],
      cell = "",
      q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            cell += '"';
            i++;
          } else q = false;
        } else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === d) {
        row.push(cell);
        cell = "";
      } else if (ch === "\n") {
        row.push(cell);
        rows.push(row);
        row = [];
        cell = "";
      } else if (ch !== "\r") cell += ch;
    }
    if (cell !== "" || row.length) {
      row.push(cell);
      rows.push(row);
    }
    return rows;
  }
  /* Abre um arquivo (.csv ou Excel) e devolve as abas com as linhas. */
  async function sheetsOf(f) {
    const buf = await f.arrayBuffer();
    if (/\.(csv|txt)$/i.test(f.name)) return [{ name: "", rows: parseCSV(decode(buf)) }];
    if (!window.XLSX)
      throw new Error("o leitor de Excel não carregou. Confira a internet e recarregue a página.");
    const wb = XLSX.read(buf, { type: "array" });
    return wb.SheetNames.map((n) => ({
      name: n,
      rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: false }),
    }));
  }
  function setLoading(t) {
    const l = $("#loading");
    if (t) {
      l.textContent = t;
      l.hidden = false;
    } else {
      l.hidden = true;
      l.textContent = "Lendo planilhas…";
    }
  }
  function showReport(list) {
    $("#report").innerHTML = list
      .map(
        (r) =>
          `<div class="rep glass ${r.ok ? "ok" : "err"}"><i class="dot"></i><span>${r.ok ? "" : `<b>${esc(r.file)}</b>: `}${r.msg}</span><button class="x" type="button" aria-label="Fechar aviso">✕</button></div>`,
      )
      .join("");
  }
  const countTipos = (txs) => {
    let r = 0,
      d = 0,
      n = 0;
    for (const t of txs) {
      if (t.tipo === "RECEITA") r++;
      else if (t.tipo === "DESPESA") d++;
      else if (t.tipo === "NP") n++;
    }
    return `${plural(r, "receita", "receitas")}, ${plural(d, "despesa", "despesas")}${n ? ` e ${plural(n, "aporte (NP)", "aportes (NP)")}` : ""}`;
  };

  /* Upload manual: lê cada arquivo e junta às planilhas já carregadas. */
  async function handleFiles(list) {
    const files = [...list].filter(
      (f) => /\.(xlsx|xlsm|xls|csv|txt|ods)$/i.test(f.name) && !f.name.startsWith("~$"),
    );
    if (!files.length) {
      showReport([
        { ok: false, file: "Arquivo não aceito", msg: "use planilhas .xlsm, .xlsx, .xls ou .csv." },
      ]);
      return;
    }
    if (state.embedded) state = { accounts: {}, updatedAt: null, embedded: false };
    const out = [];
    const batch = new Set();
    let fi = 0;
    for (const f of files) {
      setLoading(`Lendo ${++fi} de ${files.length}: ${f.name}`);
      await new Promise((r) => setTimeout(r, 0));
      try {
        const sheets = await sheetsOf(f);
        const accs = parseBook(sheets, f.name);
        if (!accs.length) {
          out.push({
            ok: false,
            file: f.name,
            msg: "não encontrei a aba de conciliação (Tipo, DATA, LANÇAMENTO, VALOR) nem um extrato com Data e Valor.",
          });
          continue;
        }
        const a = accs[0];
        const { key, prev } = resolveKey(a, batch);
        batch.add(key);
        const ds = a.txs
          .map((t) => t.date)
          .concat(a.saldos.map((x) => x.date))
          .sort();
        const mn = ds[0],
          mx = ds[ds.length - 1];
        let txs = a.txs,
          saldos = a.saldos;
        if (prev) {
          txs = prev.txs.filter((t) => t.date < mn || t.date > mx).concat(a.txs);
          saldos = prev.saldos.filter((x) => x.date < mn || x.date > mx).concat(a.saldos);
        }
        const byD = (x, y) => (x.date < y.date ? -1 : x.date > y.date ? 1 : 0);
        txs.sort(byD);
        saldos.sort(byD);
        state.accounts[key] = {
          name: key,
          pv: PARSER_V,
          srcName: a.name,
          title: a.title,
          planoHint: a.planoHint,
          file: f.name,
          sheet: a.sheet,
          tabs: a.tabs,
          format: a.format,
          txs,
          saldos,
          loadedAt: new Date().toISOString(),
        };
        const p = planoOf(state.accounts[key]);
        out.push({
          ok: true,
          msg: `<b>${esc(key)}</b> ${prev ? "atualizada" : "adicionada"}: ${countTipos(a.txs)} de ${fDay(mn)} a ${fDay(mx)} · ${p ? esc(plInfo(p).full) : '<b style="color:var(--warn)">sem plano: escolha na aba Planilhas</b>'}`,
        });
      } catch (e) {
        out.push({ ok: false, file: f.name, msg: "não consegui abrir: " + esc(e.message || e) });
      }
    }
    state.updatedAt = new Date().toISOString();
    const ok = await save();
    if (!ok)
      out.push({
        ok: false,
        file: "Aviso",
        msg: "não consegui guardar as planilhas neste navegador. Elas ficam abertas até você fechar a página.",
      });
    setLoading(null);
    rebuild();
    render();
    showReport(out);
  }
  /* Decide se o arquivo atualiza uma planilha que já existe ou entra como nova. */
  function resolveKey(a, batch) {
    for (const [k, ex] of Object.entries(state.accounts)) {
      if (ex.src === "pasta" || (ex.srcName || ex.name) !== a.name || batch.has(k)) continue;
      if (ex.title && a.title && norm(ex.title) !== norm(a.title)) continue;
      return { key: k, prev: ex };
    }
    let key = a.name;
    if (state.accounts[key]) {
      key = a.title && !state.accounts[`${a.name} · ${a.title}`] ? `${a.name} · ${a.title}` : a.name;
      let i = 2;
      while (state.accounts[key]) key = `${a.name} (${i++})`;
    }
    return { key, prev: null };
  }

  /* ---------- pasta do SharePoint (sincronizada pelo OneDrive) ---------- */
  const PASTA = {
    ok: "showDirectoryPicker" in window && !FRAMED,
    handle: null,
    perm: "",
    lendo: false,
    erros: [],
  };
  /* Pasta do computador: lista as planilhas da pasta e das subpastas (até 3 níveis). */
  async function listBooks(dir, prefix = "", depth = 0, out = []) {
    for await (const [name, h] of dir.entries()) {
      if (name.startsWith(".") || name.startsWith("~$")) continue;
      if (h.kind === "file" && /\.(xlsm|xlsx|xls|csv)$/i.test(name)) out.push([prefix + name, h]);
      else if (h.kind === "directory" && depth < 3) await listBooks(h, prefix + name + "/", depth + 1, out);
    }
    return out;
  }
  /* Lê os arquivos da pasta (ou do SharePoint). Só abre de novo o que mudou (tamanho/data). */
  async function readEntries(entries, force) {
    if (state.embedded) state = { accounts: {}, updatedAt: null, embedded: false };
    const files = state.files || (state.files = {});
    const seen = new Set(entries.map((e) => e[0]));
    const mudou = [],
      saiu = [],
      erros = [],
      ign = [];
    for (const p of Object.keys(files))
      if (!seen.has(p)) {
        if (!files[p].ignored) saiu.push(p);
        delete files[p];
      }
    const todo = entries.filter(
      ([p, f]) =>
        force ||
        !files[p] ||
        files[p].sig !== f.size + ":" + f.lastModified ||
        (!files[p].ignored && (files[p].pv || 1) < PARSER_V),
    );
    let i = 0;
    for (const [p, f] of todo) {
      setLoading(`Lendo ${++i} de ${todo.length}: ${f.name}`);
      await new Promise((r) => setTimeout(r, 0));
      const sig = f.size + ":" + f.lastModified;
      try {
        const sheets = await sheetsOf(f);
        const a = parseBook(sheets, f.name)[0] || null;
        if (!a) {
          files[p] = { sig, ignored: true, file: f.name };
          ign.push(p);
          continue;
        }
        const was = files[p] && !files[p].ignored;
        files[p] = {
          sig,
          pv: PARSER_V,
          lastModified: f.lastModified,
          file: f.name,
          name: a.name,
          title: a.title,
          planoHint: planoFromName(p) || a.planoHint,
          sheet: a.sheet,
          tabs: a.tabs,
          format: a.format,
          txs: a.txs,
          saldos: a.saldos,
        };
        mudou.push({ p, novo: !was });
      } catch (e) {
        erros.push(
          `${p}: não consegui abrir (${e.message || e}). Se o arquivo está só na nuvem, espere o OneDrive terminar de baixar.`,
        );
      }
    }
    for (const p of Object.keys(files)) if (files[p].ignored && !ign.includes(p)) ign.push(p);
    const replaced = compose();
    return {
      mudou,
      saiu,
      erros,
      ign,
      replaced,
      total: Object.values(files).filter((f) => !f.ignored).length,
    };
  }
  /* arquivos da pasta com o mesmo nome (sem data, mês ou trimestre) viram uma planilha só */
  /* Junta os arquivos da pasta em planilhas: arquivos com o mesmo nome (ex.: um por trimestre) viram uma só. */
  function compose() {
    const files = state.files || {};
    const groups = new Map();
    for (const [p, f] of Object.entries(files)) {
      if (f.ignored || !f.txs || !f.txs.length) continue;
      const gk = norm(f.name) + "|" + norm(f.title || "");
      if (!groups.has(gk)) groups.set(gk, []);
      groups.get(gk).push([p, f]);
    }
    for (const k of Object.keys(state.accounts))
      if (state.accounts[k].src === "pasta") delete state.accounts[k];
    const cnt = {};
    for (const l of groups.values()) cnt[l[0][1].name] = (cnt[l[0][1].name] || 0) + 1;
    const replaced = [];
    const byD = (x, y) => (x.date < y.date ? -1 : x.date > y.date ? 1 : 0);
    for (const [gk, list] of groups) {
      list.sort((a, b) => (a[1].lastModified || 0) - (b[1].lastModified || 0));
      const f0 = list[0][1],
        last = list[list.length - 1][1];
      let name =
        (state.alias || {})[gk] || (cnt[f0.name] > 1 && f0.title ? `${f0.name} · ${f0.title}` : f0.name);
      for (const [k, ac] of Object.entries(state.accounts))
        if (ac.src !== "pasta" && ((ac.srcName || ac.name) === f0.name || k === name)) {
          delete state.accounts[k];
          replaced.push(k);
        }
      const base = name;
      let i = 2;
      while (state.accounts[name]) name = `${base} (${i++})`;
      let txs = [],
        saldos = [];
      for (const [, f] of list) {
        const ds = f.txs
          .map((t) => t.date)
          .concat(f.saldos.map((x) => x.date))
          .sort();
        const mn = ds[0],
          mx = ds[ds.length - 1];
        txs = txs.filter((t) => t.date < mn || t.date > mx).concat(f.txs);
        saldos = saldos.filter((x) => x.date < mn || x.date > mx).concat(f.saldos);
      }
      txs.sort(byD);
      saldos.sort(byD);
      const hints = {};
      for (const [p, f] of list) {
        const h = f.planoHint || planoFromName(p);
        if (h) hints[h] = (hints[h] || 0) + 1;
      }
      const hint = (Object.entries(hints).sort((a, b) => b[1] - a[1])[0] || [""])[0];
      state.accounts[name] = {
        name,
        pv: Math.min(...list.map((x) => x[1].pv || 1)),
        srcName: f0.name,
        title: f0.title,
        src: "pasta",
        gk,
        paths: list.map((x) => x[0]),
        file: list.map((x) => x[0]).join(", "),
        planoHint: hint,
        sheet: last.sheet,
        tabs: last.tabs,
        format: last.format,
        txs,
        saldos,
      };
    }
    return replaced;
  }
  function folderReport(r, force) {
    const out = [];
    const nm = state.folderName || "pasta";
    if (force || r.mudou.length || r.saiu.length) {
      const novas = r.mudou.filter((x) => x.novo).length,
        alt = r.mudou.length - novas;
      const parts = [];
      if (force) parts.push(plural(r.total, "arquivo lido", "arquivos lidos"));
      else {
        if (novas) parts.push(plural(novas, "arquivo novo", "arquivos novos"));
        if (alt) parts.push(plural(alt, "arquivo alterado", "arquivos alterados"));
      }
      if (r.saiu.length) parts.push(`${r.saiu.length} ${r.saiu.length === 1 ? "saiu" : "saíram"} da pasta`);
      const quais = r.mudou.map((x) => x.p.split("/").pop());
      out.push({
        ok: true,
        msg: `<b>${esc(nm)}</b> atualizada ${stamp(new Date())}: ${parts.join(" · ")}${quais.length && !force ? ` (${esc(quais.slice(0, 6).join(", "))}${quais.length > 6 ? "…" : ""})` : ""}.`,
      });
    }
    if (r.replaced.length)
      out.push({
        ok: true,
        msg: `Tirei ${esc(r.replaced.join(", "))}, que tinha sido carregada à mão e agora vem da pasta.`,
      });
    if (force && r.ign.length)
      out.push({
        ok: true,
        msg: `${r.ign.length === 1 ? "1 arquivo da pasta não tem" : `${r.ign.length} arquivos da pasta não têm`} aba de conciliação e ${r.ign.length === 1 ? "foi ignorado" : "foram ignorados"}: ${esc(
          r.ign
            .slice(0, 5)
            .map((p) => p.split("/").pop())
            .join(", "),
        )}${r.ign.length > 5 ? "…" : ""}.`,
      });
    const miss = ORDER.filter((n) => !planoOf(state.accounts[n])).length;
    if (miss)
      out.push({
        ok: false,
        file: "Plano",
        msg: `${plural(miss, "planilha está", "planilhas estão")} sem plano. Escolha na aba Planilhas.`,
      });
    for (const e of r.erros) out.push({ ok: false, file: "Pasta", msg: esc(e) });
    return out;
  }
  /* Relê a pasta do computador (force = abre todos os arquivos de novo). */
  async function syncFolder(force) {
    if (!PASTA.handle || PASTA.lendo || SP.cfg.active) return null;
    let r = null;
    try {
      PASTA.perm = await PASTA.handle.queryPermission({ mode: "read" });
      if (PASTA.perm !== "granted") {
        renderFolder();
        return null;
      }
      PASTA.lendo = true;
      renderFolder();
      if (force) setLoading("Procurando planilhas na pasta…");
      const list = await listBooks(PASTA.handle);
      const entries = await Promise.all(list.map(async ([p, h]) => [p, await h.getFile()]));
      r = await readEntries(entries, force);
      PASTA.erros = r.erros.slice();
      if (!r.total && !r.ign.length)
        PASTA.erros = ["nenhuma planilha (.xlsm, .xlsx, .xls ou .csv) nessa pasta nem nas subpastas."];
      else if (!r.total)
        PASTA.erros = [
          "nenhuma planilha dessa pasta tem a aba de conciliação (Tipo, DATA, LANÇAMENTO, VALOR).",
        ];
      state.folderName = PASTA.handle.name;
      state.folderRead = new Date().toISOString();
      state.srcKind = "local";
      if (force || r.mudou.length || r.saiu.length || r.replaced.length) {
        state.updatedAt = state.folderRead;
        await save();
        rebuild();
        render();
        showReport(folderReport(r, force));
      } else await save();
    } catch (e) {
      PASTA.erros = ["não consegui ler a pasta: " + (e.message || e)];
    } finally {
      PASTA.lendo = false;
      setLoading(null);
      renderFolder();
      $("#btnSync").disabled = false;
    }
    return r;
  }
  /* Escolher/trocar a pasta do computador (Edge/Chrome guardam a escolha). */
  async function connectFolder() {
    if (!PASTA.ok) {
      $("#foldIn").click();
      return;
    }
    try {
      const h = await window.showDirectoryPicker({ id: "fluxo-caixa", mode: "read" });
      const trocou = PASTA.handle && !(await PASTA.handle.isSameEntry(h).catch(() => false));
      PASTA.handle = h;
      PASTA.perm = "granted";
      try {
        await IDB.set("pasta", h);
      } catch (e) {}
      if (trocou || state.srcKind === "sp") {
        state.files = {};
        state.tabPref = {};
      }
      if (SP.cfg.active) {
        SP.cfg.active = false;
        spSaveCfg();
      }
      await syncFolder(true);
    } catch (e) {
      if (e.name !== "AbortError") {
        PASTA.erros = [
          e.name === "SecurityError"
            ? "o navegador não deixou abrir a pasta aqui. Abra o arquivo fluxo-caixa.html direto no Edge ou no Chrome."
            : e.message || String(e),
        ];
        renderFolder();
      }
    }
  }
  async function allowFolder() {
    if (!PASTA.handle) return connectFolder();
    try {
      PASTA.perm = await PASTA.handle.requestPermission({ mode: "read" });
    } catch (e) {}
    if (PASTA.perm === "granted") await syncFolder(true);
    else renderFolder();
  }
  async function updateAll() {
    if (SP.cfg.active) return spUpdate();
    if (PASTA.handle) return PASTA.perm === "granted" ? syncFolder(true) : allowFolder();
    if (PASTA.ok) return connectFolder();
    $("#foldIn").click();
  }
  async function forgetFolder() {
    try {
      await IDB.set("pasta", null);
    } catch (e) {}
    PASTA.handle = null;
    PASTA.perm = "";
    PASTA.erros = [];
    state.files = {};
    state.folderName = "";
    state.folderRead = null;
    compose();
    await save();
    rebuild();
    render();
    showReport([{ ok: true, msg: "Pasta desconectada. As planilhas que vinham dela saíram do painel." }]);
  }
  async function folderFromInput(list) {
    const entries = [...list]
      .filter((f) => /\.(xlsm|xlsx|xls|csv)$/i.test(f.name) && !f.name.startsWith("~$"))
      .map((f) => [(f.webkitRelativePath || f.name).split("/").slice(1).join("/") || f.name, f]);
    const top = ((list[0] && list[0].webkitRelativePath) || "").split("/")[0];
    if (!entries.length) {
      PASTA.erros = ["nenhuma planilha (.xlsm, .xlsx, .xls ou .csv) nessa pasta."];
      renderFolder();
      return;
    }
    const r = await readEntries(entries, true);
    setLoading(null);
    PASTA.erros = r.erros;
    state.folderName = top || "pasta";
    state.folderRead = new Date().toISOString();
    state.updatedAt = state.folderRead;
    await save();
    rebuild();
    render();
    showReport(folderReport(r, true));
  }
  /* ---------- SharePoint pelo Microsoft Graph ----------
   Login com a conta do trabalho (OAuth 2.0 com PKCE, como a biblioteca oficial da Microsoft faz),
   leitura da pasta e das subpastas pelo Graph e atualização automática das planilhas que mudarem.
   Para todos já abrirem configurado, o TI pode preencher SP_DEFAULT aqui embaixo. */
  const SP_DEFAULT = { clientId: "", tenant: "", url: "" };
  const SP_SCOPES = "openid profile offline_access Files.Read.All Sites.Read.All";
  const SP = {
    ok: !FRAMED && /^https?:$/.test(location.protocol) && !!(window.crypto && crypto.subtle),
    cfg: { ...SP_DEFAULT, active: false, resolved: null },
    tok: null,
    user: "",
    lendo: false,
    erros: [],
    edit: false,
  };
  try {
    const c = JSON.parse(localStorage.getItem(STORE + "-sp") || "null");
    if (c) SP.cfg = { ...SP.cfg, ...c };
  } catch (e) {}
  try {
    const t = JSON.parse(localStorage.getItem(STORE + "-sptok") || "null");
    if (t && t.rt && t.rtExp > Date.now()) {
      SP.tok = t;
      SP.user = t.user || "";
    }
  } catch (e) {}
  function spSaveCfg() {
    try {
      localStorage.setItem(STORE + "-sp", JSON.stringify(SP.cfg));
    } catch (e) {}
  }
  function spSaveTok() {
    try {
      if (SP.tok) localStorage.setItem(STORE + "-sptok", JSON.stringify(SP.tok));
      else localStorage.removeItem(STORE + "-sptok");
    } catch (e) {}
  }
  const spRedirect = () => location.origin + location.pathname;
  const spAuthBase = () =>
    `https://login.microsoftonline.com/${encodeURIComponent((SP.cfg.tenant || "organizations").trim())}/oauth2/v2.0`;
  const b64url = (buf) => {
    let s = "";
    const b = new Uint8Array(buf);
    for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]);
    return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  };
  const rand = (n) => b64url(crypto.getRandomValues(new Uint8Array(n)));
  function jwtUser(idt) {
    try {
      const p = JSON.parse(
        decodeURIComponent(escape(atob(idt.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")))),
      );
      return p.preferred_username || p.upn || p.email || p.name || "";
    } catch (e) {
      return "";
    }
  }
  function spErr(msg, status) {
    const m = String(msg || "");
    let t = m;
    if (/AADSTS50011/.test(m))
      t = `o endereço desta página não está cadastrado no aplicativo do Azure. Peça ao TI para incluir ${spRedirect()} como URI de redirecionamento do tipo "Aplicativo de página única (SPA)".`;
    else if (/AADSTS9002326/.test(m))
      t =
        'o endereço está cadastrado como "Web" no Azure. Ele precisa estar como "Aplicativo de página única (SPA)".';
    else if (/AADSTS65001|AADSTS90094|consent/i.test(m))
      t =
        "o aplicativo ainda não tem as permissões aprovadas. Peça ao TI o consentimento de administrador para Files.Read.All e Sites.Read.All.";
    else if (/AADSTS700016/.test(m))
      t = "o ID do aplicativo não existe nesse locatário. Confira o ID do aplicativo e o locatário.";
    else if (/AADSTS90002|AADSTS900023/.test(m))
      t = "o locatário não foi encontrado. Use o ID do locatário (Diretório) ou o domínio da empresa.";
    else if (/AADSTS70000|AADSTS700082|invalid_grant/i.test(m))
      t = "o login expirou. Clique em Entrar de novo.";
    else if (/access_denied|AADSTS65004/i.test(m)) t = "o login foi cancelado.";
    const e = new Error(t);
    e.status = status;
    return e;
  }
  /* SharePoint: pede o token de acesso ao login da Microsoft (código de login ou renovação). */
  async function spTokenReq(params) {
    let r,
      j = {};
    try {
      r = await fetch(spAuthBase() + "/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ client_id: SP.cfg.clientId.trim(), scope: SP_SCOPES, ...params }),
      });
      j = await r.json().catch(() => ({}));
    } catch (e) {
      throw new Error(
        "não consegui falar com o login da Microsoft. Confira a internet; se continuar, a rede pode estar bloqueando login.microsoftonline.com.",
      );
    }
    if (!r.ok || !j.access_token) throw spErr(j.error_description || j.error || "erro " + r.status);
    const old = SP.tok || {};
    SP.tok = {
      at: j.access_token,
      atExp: Date.now() + ((j.expires_in || 3600) - 120) * 1000,
      rt: j.refresh_token || old.rt,
      rtExp: j.refresh_token ? Date.now() + 23 * 3600e3 : old.rtExp || 0,
      user: j.id_token ? jwtUser(j.id_token) : old.user || "",
    };
    SP.user = SP.tok.user;
    spSaveTok();
    return SP.tok.at;
  }
  /* SharePoint: devolve um token válido, renovando quando vence; null se precisar entrar de novo. */
  async function spAccess() {
    if (SP.tok && SP.tok.at && SP.tok.atExp > Date.now()) return SP.tok.at;
    if (SP.tok && SP.tok.rt && SP.tok.rtExp > Date.now()) {
      try {
        return await spTokenReq({ grant_type: "refresh_token", refresh_token: SP.tok.rt });
      } catch (e) {
        SP.tok = null;
        spSaveTok();
        return null;
      }
    }
    return null;
  }
  /* login numa janela pop-up; a página volta com ?code=… e entrega pelo postMessage ou pelo localStorage */
  /* SharePoint: login numa janela pop-up (OAuth 2.0 com PKCE). */
  function spLogin() {
    return new Promise((resolve, reject) => {
      if (!SP.cfg.clientId.trim()) {
        reject(new Error("preencha o ID do aplicativo."));
        return;
      }
      const verifier = rand(48),
        state = rand(16);
      try {
        localStorage.removeItem("fc-oauth-ret");
      } catch (e) {}
      const w = window.open("about:blank", "fc-login", "width=520,height=680");
      if (!w) {
        reject(
          new Error(
            "o navegador bloqueou a janela de login. Permita pop-ups para este endereço e clique de novo.",
          ),
        );
        return;
      }
      crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)).then((h) => {
        w.location.href =
          spAuthBase() +
          "/authorize?" +
          new URLSearchParams({
            client_id: SP.cfg.clientId.trim(),
            response_type: "code",
            redirect_uri: spRedirect(),
            response_mode: "query",
            scope: SP_SCOPES,
            state,
            code_challenge: b64url(h),
            code_challenge_method: "S256",
            prompt: "select_account",
          });
      });
      let done = false,
        closedAt = 0;
      const finish = async (ret) => {
        if (done) return;
        done = true;
        clearInterval(poll);
        window.removeEventListener("message", onMsg);
        window.removeEventListener("storage", onSto);
        try {
          w.close();
        } catch (e) {}
        try {
          localStorage.removeItem("fc-oauth-ret");
        } catch (e) {}
        if (!ret) {
          reject(new Error("a janela de login foi fechada antes de terminar."));
          return;
        }
        if (ret.state !== state) {
          reject(new Error("o login não confere com este pedido. Tente de novo."));
          return;
        }
        if (ret.error) {
          reject(spErr(ret.desc || ret.error));
          return;
        }
        try {
          await spTokenReq({
            grant_type: "authorization_code",
            code: ret.code,
            redirect_uri: spRedirect(),
            code_verifier: verifier,
          });
          resolve();
        } catch (e) {
          reject(e);
        }
      };
      const onMsg = (e) => {
        if (e.origin === location.origin && e.data && e.data.fcAuth) finish(e.data);
      };
      const onSto = (e) => {
        if (e.key === "fc-oauth-ret" && e.newValue) {
          try {
            finish(JSON.parse(e.newValue));
          } catch (x) {}
        }
      };
      window.addEventListener("message", onMsg);
      window.addEventListener("storage", onSto);
      const poll = setInterval(() => {
        let r = null;
        try {
          r = JSON.parse(localStorage.getItem("fc-oauth-ret") || "null");
        } catch (e) {}
        if (r) {
          finish(r);
          return;
        }
        try {
          if (w.location.origin === location.origin) {
            const q = new URLSearchParams(w.location.search);
            if (q.get("state")) {
              finish({
                code: q.get("code"),
                error: q.get("error"),
                desc: q.get("error_description"),
                state: q.get("state"),
              });
              return;
            }
          }
        } catch (e) {}
        if (w.closed) {
          if (!closedAt) closedAt = Date.now();
          else if (Date.now() - closedAt > 1500) finish(null);
        }
      }, 350);
    });
  }
  /* SharePoint: chamada ao Microsoft Graph com o token, com mensagens de erro em português. */
  async function g(path, retry) {
    const at = await spAccess();
    if (!at) {
      const e = new Error("entre com a conta Microsoft para ler o SharePoint.");
      e.login = true;
      throw e;
    }
    let r;
    try {
      r = await fetch(path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path, {
        headers: { Authorization: "Bearer " + at },
      });
    } catch (e) {
      throw new Error(
        "não consegui falar com o Microsoft Graph. Confira a internet; a rede pode estar bloqueando graph.microsoft.com.",
      );
    }
    if (r.status === 401 && !retry && SP.tok) {
      SP.tok.atExp = 0;
      return g(path, true);
    }
    if (!r.ok) {
      const j = await r.json().catch(() => ({}));
      const msg = (j.error && j.error.message) || "";
      throw spErr(
        r.status === 403
          ? "sua conta não tem acesso a essa pasta no SharePoint, ou o aplicativo não tem a permissão Sites.Read.All aprovada."
          : r.status === 404
            ? "não encontrei essa pasta no SharePoint. Confira o link."
            : r.status === 401
              ? "o login expirou. Clique em Entrar de novo."
              : msg || "erro " + r.status,
        r.status,
      );
    }
    return r.json();
  }
  /* aceita o link da pasta como aparece no navegador (inclusive AllItems.aspx?id=…) ou um link de compartilhamento */
  /* SharePoint: limpa o link colado (aceita o endereço AllItems.aspx?id=… ou link de compartilhamento). */
  function spNormUrl(raw) {
    let u = String(raw || "").trim();
    if (!u) return "";
    try {
      const x = new URL(u);
      const id = x.searchParams.get("id") || x.searchParams.get("RootFolder");
      if (id && /\.aspx$/i.test(x.pathname)) return x.origin + decodeURIComponent(id);
      if (/\/:[a-z]:\//i.test(x.pathname)) return u; /* link de compartilhamento */
      return x.origin + decodeURIComponent(x.pathname).replace(/\/+$/, "");
    } catch (e) {
      return u;
    }
  }
  const spShareToken = (u) =>
    "u!" +
    btoa(unescape(encodeURIComponent(u)))
      .replace(/=+$/, "")
      .replace(/\//g, "_")
      .replace(/\+/g, "-");
  /* SharePoint: encontra a pasta pelo link (primeiro como link de compartilhamento; senão site › biblioteca › caminho). */
  async function spResolve(raw) {
    const u = spNormUrl(raw);
    let x;
    try {
      x = new URL(u);
    } catch (e) {
      throw new Error(
        "o link da pasta não parece um endereço do SharePoint. Copie o endereço da pasta na barra do navegador.",
      );
    }
    if (!/sharepoint\.com$/i.test(x.hostname))
      throw new Error("o link precisa ser de um endereço sharepoint.com.");
    try {
      const it = await g(`/shares/${spShareToken(u)}/driveItem`);
      if (it.folder)
        return { driveId: it.parentReference.driveId, id: it.id, name: it.name, webUrl: it.webUrl };
    } catch (e) {
      if (e.login) throw e;
    }
    const segs = decodeURIComponent(x.pathname).split("/").filter(Boolean);
    let sitePath = "",
      rest = segs;
    if (/^(sites|teams)$/i.test(segs[0] || "")) {
      sitePath = segs.slice(0, 2).join("/");
      rest = segs.slice(2);
    }
    const site = await g(
      sitePath
        ? `/sites/${x.hostname}:/${sitePath.split("/").map(encodeURIComponent).join("/")}`
        : `/sites/${x.hostname}`,
    );
    const drives = (await g(`/sites/${site.id}/drives`)).value || [];
    for (const d of drives) {
      let dl = "";
      try {
        const ds = decodeURIComponent(new URL(d.webUrl).pathname).split("/").filter(Boolean);
        dl = ds[ds.length - 1] || "";
      } catch (e) {}
      if (!rest.length || (norm(dl) !== norm(rest[0]) && norm(d.name) !== norm(rest[0]))) continue;
      const rel = rest.slice(1);
      const it = await g(
        rel.length
          ? `/drives/${d.id}/root:/${rel.map(encodeURIComponent).join("/")}`
          : `/drives/${d.id}/root`,
      );
      if (!it.folder) throw new Error("o link aponta para um arquivo. Cole o link da pasta.");
      return { driveId: d.id, id: it.id, name: it.name === "root" ? d.name : it.name, webUrl: it.webUrl };
    }
    throw new Error(
      "não encontrei a biblioteca de documentos desse link. Copie o endereço da pasta direto da barra do navegador.",
    );
  }
  /* SharePoint: lista as planilhas da pasta e das subpastas, com o endereço de download de cada uma. */
  async function spList(driveId, id, prefix = "", depth = 0, out = []) {
    let next = `/drives/${driveId}/items/${id}/children?$top=200`;
    while (next) {
      const j = await g(next);
      for (const it of j.value || []) {
        if (it.name.startsWith("~$") || it.name.startsWith(".")) continue;
        if (it.folder) {
          if (depth < 3) await spList(driveId, it.id, prefix + it.name + "/", depth + 1, out);
          continue;
        }
        if (!it.file || !/\.(xlsm|xlsx|xls|csv)$/i.test(it.name)) continue;
        const dl = it["@microsoft.graph.downloadUrl"];
        out.push([
          prefix + it.name,
          {
            name: it.name,
            size: it.size || 0,
            lastModified: Date.parse(it.lastModifiedDateTime) || 0,
            arrayBuffer: async () => {
              let r;
              if (dl) r = await fetch(dl);
              else {
                const at = await spAccess();
                r = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${it.id}/content`, {
                  headers: { Authorization: "Bearer " + at },
                });
              }
              if (!r.ok) throw new Error("não consegui baixar o arquivo (" + r.status + ")");
              return r.arrayBuffer();
            },
          },
        ]);
      }
      next = j["@odata.nextLink"] || null;
    }
    return out;
  }
  /* SharePoint: lê a pasta e atualiza o painel (force = baixa tudo de novo). */
  async function syncSP(force) {
    if (!SP.ok || !SP.cfg.active || SP.lendo) return null;
    let r = null;
    SP.lendo = true;
    renderSP();
    try {
      if (!(await spAccess())) {
        SP.erros = [];
        return null;
      }
      if (force) setLoading("Conectando ao SharePoint…");
      if (!SP.cfg.resolved || SP.cfg.resolved.src !== SP.cfg.url) {
        SP.cfg.resolved = { ...(await spResolve(SP.cfg.url)), src: SP.cfg.url };
        spSaveCfg();
      }
      const R = SP.cfg.resolved;
      if (force) setLoading("Procurando planilhas no SharePoint…");
      const entries = await spList(R.driveId, R.id);
      if (state.srcKind !== "sp") {
        state.files = {};
        state.tabPref = {};
      }
      r = await readEntries(entries, force);
      SP.erros = r.erros.slice();
      if (!r.total && !r.ign.length)
        SP.erros = ["nenhuma planilha (.xlsm, .xlsx, .xls ou .csv) nessa pasta nem nas subpastas."];
      else if (!r.total)
        SP.erros = ["nenhuma planilha dessa pasta tem a aba de conciliação (Tipo, DATA, LANÇAMENTO, VALOR)."];
      state.srcKind = "sp";
      state.folderName = R.name;
      state.folderRead = new Date().toISOString();
      if (force || r.mudou.length || r.saiu.length || r.replaced.length) {
        state.updatedAt = state.folderRead;
        await save();
        rebuild();
        render();
        showReport(folderReport(r, force));
      } else await save();
    } catch (e) {
      SP.erros = [e.message || String(e)];
      if (e.status === 404) {
        SP.cfg.resolved = null;
        spSaveCfg();
      }
    } finally {
      SP.lendo = false;
      setLoading(null);
      renderSP();
      renderFolder();
    }
    return r;
  }
  function spReadForm() {
    const v = (id) => {
      const el = document.getElementById(id);
      return el ? el.value.trim() : null;
    };
    const c = v("spClient"),
      t = v("spTenant"),
      u = v("spUrl");
    if (c != null) SP.cfg.clientId = c;
    if (t != null) SP.cfg.tenant = t;
    if (u != null && u !== SP.cfg.url) {
      SP.cfg.url = u;
      SP.cfg.resolved = null;
    }
    spSaveCfg();
  }
  /* SharePoint: valida os campos, faz o login e conecta. */
  async function spConnect() {
    spReadForm();
    SP.erros = [];
    const miss = [
      !SP.cfg.clientId && "o ID do aplicativo",
      !SP.cfg.tenant && "o locatário",
      !SP.cfg.url && "o link da pasta",
    ].filter(Boolean);
    if (miss.length) {
      SP.erros = [`preencha ${miss.join(", ")}.`];
      renderSP();
      return;
    }
    try {
      if (!(await spAccess())) {
        SP.lendo = true;
        renderSP();
        await spLogin();
      }
    } catch (e) {
      SP.lendo = false;
      SP.erros = [e.message || String(e)];
      renderSP();
      return;
    }
    SP.lendo = false;
    if (state.srcKind !== "sp") {
      state.files = {};
      state.tabPref = {};
    }
    SP.cfg.active = true;
    SP.edit = false;
    spSaveCfg();
    await syncSP(true);
  }
  async function spUpdate() {
    if (await spAccess()) return syncSP(true);
    return spConnect();
  }
  function spLogout() {
    SP.tok = null;
    SP.user = "";
    spSaveTok();
    SP.erros = [];
    renderSP();
  }
  async function spDisconnect() {
    SP.cfg.active = false;
    SP.cfg.resolved = null;
    spSaveCfg();
    SP.tok = null;
    spSaveTok();
    SP.erros = [];
    if (state.srcKind === "sp") {
      state.files = {};
      state.folderName = "";
      state.folderRead = null;
      state.srcKind = "";
      compose();
      await save();
      rebuild();
    }
    render();
    showReport([
      { ok: true, msg: "SharePoint desconectado. As planilhas que vinham de lá saíram do painel." },
    ]);
  }
  function spCardHTML() {
    return `<section class="glass folder" id="spCard"><span class="fic" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7 18a4.5 4.5 0 0 1-.6-8.96A6 6 0 0 1 18 8.5a4 4 0 0 1-1 9.5z"/></svg></span>
    <div class="fbody"><div class="ftop"><b>SharePoint (Microsoft Graph)</b><span id="spStat"></span></div><div class="finfo" id="spInfo"></div></div><div class="acts" id="spActs"></div>
    <div id="spForm" style="flex:1 1 100%"></div><div class="ferr" id="spErr"></div></section>`;
  }
  /* SharePoint: desenha o card de conexão na aba Planilhas. */
  function renderSP() {
    $("#btnSync").hidden = !(PASTA.handle || state.folderName || SP.cfg.active);
    $("#btnSync").disabled = PASTA.lendo || SP.lendo;
    const box = $("#spCard");
    if (!box) return;
    const st = $("#spStat"),
      info = $("#spInfo"),
      acts = $("#spActs"),
      form = $("#spForm");
    const logged = !!(
      SP.tok &&
      ((SP.tok.at && SP.tok.atExp > Date.now()) || (SP.tok.rt && SP.tok.rtExp > Date.now()))
    );
    const n = state.srcKind === "sp" ? Object.values(state.files || {}).filter((f) => !f.ignored).length : 0;
    const lida =
      state.srcKind === "sp" && state.folderRead ? ` · última leitura ${stamp(state.folderRead)}` : "";
    box.className = "glass folder";
    form.innerHTML = "";
    if (!SP.ok) {
      st.innerHTML = '<span class="pill off">Indisponível aqui</span>';
      info.innerHTML = FRAMED
        ? "Dentro do Claude a conexão com a Microsoft é bloqueada. Abra o painel pelo endereço onde ele está publicado (Vercel ou SharePoint) para entrar com a conta do trabalho."
        : "Aberto como arquivo no computador, o login da Microsoft não funciona. Publique o painel num endereço https (por exemplo no Vercel), cadastre esse endereço no aplicativo do Azure e abra por lá.";
      acts.innerHTML = "";
    } else if (SP.cfg.active && !SP.edit) {
      box.classList.add(logged ? "on" : "wait");
      st.innerHTML = SP.lendo
        ? '<span class="pill ok spin">Lendo…</span>'
        : logged
          ? '<span class="pill ok">Conectado</span>'
          : '<span class="pill w">Falta entrar</span>';
      info.innerHTML = logged
        ? `${SP.user ? `Conta <b>${esc(SP.user)}</b> · ` : ""}lendo <b>${esc((SP.cfg.resolved && SP.cfg.resolved.name) || state.folderName || "pasta")}</b> · ${plural(n, "arquivo", "arquivos")}${lida}. O painel confere a pasta de novo quando você volta para ele e a cada 5 minutos.`
        : `A pasta <b>${esc((SP.cfg.resolved && SP.cfg.resolved.name) || state.folderName || "do SharePoint")}</b> está configurada${lida ? `, com os dados da${lida.replace(" · ", " ")}` : ""}. Entre com a conta do trabalho para atualizar.`;
      acts.innerHTML = logged
        ? `<button class="btn acc" type="button" data-sp="sync" ${SP.lendo ? "disabled" : ""}>Atualizar tudo</button><button class="btn" type="button" data-sp="edit">Configurar</button><button class="btn ghost" type="button" data-sp="logout">Sair</button><button class="btn ghost" type="button" data-sp="off">Desconectar</button>`
        : `<button class="btn acc" type="button" data-sp="connect" ${SP.lendo ? "disabled" : ""}>Entrar e atualizar</button><button class="btn" type="button" data-sp="edit">Configurar</button><button class="btn ghost" type="button" data-sp="off">Desconectar</button>`;
    } else {
      st.innerHTML = SP.lendo
        ? '<span class="pill ok spin">Entrando…</span>'
        : '<span class="pill off">Não conectado</span>';
      info.innerHTML =
        "Lê as planilhas direto da pasta do SharePoint com a sua conta do trabalho, sem precisar do OneDrive no computador. O TI cadastra um aplicativo no Azure e passa o ID do aplicativo e o locatário.";
      acts.innerHTML = `<button class="btn acc" type="button" data-sp="connect" ${SP.lendo ? "disabled" : ""}>Entrar e conectar</button>${SP.cfg.active ? '<button class="btn ghost" type="button" data-sp="cancel">Cancelar</button>' : ""}`;
      const fld = (id, lab, val, ph, hint) =>
        `<label style="display:grid;gap:4px;font-size:12.5px;color:var(--muted);min-width:0">${lab}<input id="${id}" value="${esc(val || "")}" placeholder="${esc(ph)}" autocomplete="off" spellcheck="false" style="background:rgba(0,0,0,.25);border:1px solid var(--line-2);border-radius:10px;padding:8px 11px;font-size:13px;color:var(--ink);font-family:var(--f-mono);min-width:0;width:100%">${hint ? `<span style="font-size:11.5px;color:var(--faint)">${hint}</span>` : ""}</label>`;
      form.innerHTML = `<div style="display:grid;gap:12px;grid-template-columns:repeat(auto-fit,minmax(min(100%,260px),1fr));margin-top:4px">
      ${fld("spClient", "ID do aplicativo (cliente)", SP.cfg.clientId, "00000000-0000-0000-0000-000000000000", "Azure › Registros de aplicativo › Visão geral")}
      ${fld("spTenant", "Locatário (ID do diretório ou domínio)", SP.cfg.tenant, "00000000-0000-… ou empresa.onmicrosoft.com", "Na mesma tela do aplicativo")}
      ${fld("spUrl", "Link da pasta no SharePoint", SP.cfg.url, "https://….sharepoint.com/sites/…/Bancos/2026", "Abra a pasta Bancos › 2026 no navegador e copie o endereço")}</div>
      <div style="margin-top:12px;font-size:12.5px;color:var(--muted);display:flex;flex-wrap:wrap;gap:8px;align-items:center">Endereço para o TI cadastrar no Azure (URI de redirecionamento, tipo SPA): <code style="font-family:var(--f-mono);color:var(--ink);background:rgba(0,0,0,.3);border:1px solid var(--line);border-radius:8px;padding:3px 8px;overflow-wrap:anywhere">${esc(spRedirect())}</code><button class="btn sm" type="button" data-sp="copyuri">Copiar</button></div>`;
    }
    $("#spErr").innerHTML = SP.erros.map((e) => `<span>${esc(e)}</span>`).join("");
  }
  function renderFolder() {
    const box = $("#folder");
    $("#btnSync").hidden = !(PASTA.handle || state.folderName || SP.cfg.active);
    $("#btnSync").disabled = PASTA.lendo || SP.lendo;
    if (!box) return;
    const st = $("#fstat"),
      info = $("#finfo"),
      acts = $("#facts");
    const h = PASTA.handle;
    const n = Object.values(state.files || {}).filter((f) => !f.ignored).length;
    const lida = state.folderRead ? ` · última leitura ${stamp(state.folderRead)}` : "";
    box.className = "glass folder";
    if (SP.cfg.active) {
      st.innerHTML = '<span class="pill off">Em espera</span>';
      info.innerHTML =
        "As planilhas estão vindo do SharePoint. Para usar a pasta sincronizada no computador, conecte-a aqui (o SharePoint é desligado).";
      acts.innerHTML = PASTA.ok
        ? '<button class="btn" type="button" data-fold="pick">Usar pasta do computador</button>'
        : '<button class="btn" type="button" data-fold="input">Escolher pasta</button>';
    } else if (h && PASTA.perm === "granted") {
      box.classList.add("on");
      st.innerHTML = PASTA.lendo
        ? '<span class="pill ok spin">Lendo…</span>'
        : '<span class="pill ok">Conectada</span>';
      info.innerHTML = `Lendo <b>${esc(h.name)}</b> · ${plural(n, "arquivo", "arquivos")}${lida}. O painel confere a pasta de novo quando você volta para ele e a cada 5 minutos.`;
      acts.innerHTML = `<button class="btn acc" type="button" data-fold="sync" ${PASTA.lendo ? "disabled" : ""}>Atualizar tudo</button><button class="btn" type="button" data-fold="pick">Trocar pasta</button><button class="btn ghost" type="button" data-fold="forget">Desconectar</button>`;
    } else if (h) {
      box.classList.add("wait");
      st.innerHTML = '<span class="pill w">Falta permitir</span>';
      info.innerHTML = `A pasta <b>${esc(h.name)}</b> está guardada${state.folderRead ? `, com os dados da leitura de ${stamp(state.folderRead)}` : ""}. Por segurança, o navegador pede para confirmar o acesso cada vez que você abre o painel.`;
      acts.innerHTML = `<button class="btn acc" type="button" data-fold="allow">Permitir e atualizar</button><button class="btn" type="button" data-fold="pick">Trocar pasta</button>`;
    } else if (PASTA.ok) {
      st.innerHTML = '<span class="pill off">Não conectada</span>';
      info.innerHTML =
        "Escolha a pasta do SharePoint sincronizada no computador pelo OneDrive, onde ficam as planilhas dos bancos. O painel lê todas de uma vez e atualiza sozinho quando uma muda.";
      acts.innerHTML = '<button class="btn acc" type="button" data-fold="pick">Conectar pasta</button>';
    } else {
      st.innerHTML = state.folderName
        ? '<span class="pill ok">Lida</span>'
        : '<span class="pill off">Não conectada</span>';
      info.innerHTML =
        (state.folderName
          ? `Dados da pasta <b>${esc(state.folderName)}</b> · ${plural(n, "arquivo", "arquivos")}${lida}. `
          : "") +
        (FRAMED
          ? "Para o painel lembrar da pasta e atualizar sozinho, abra o arquivo fluxo-caixa.html no Edge ou no Chrome. Aqui dá para escolher a pasta a cada vez."
          : "Este navegador não guarda a pasta. Use o Edge ou o Chrome para conectar uma vez só. Aqui dá para escolher a pasta a cada vez.");
      acts.innerHTML = `<button class="btn acc" type="button" data-fold="input">${state.folderName ? "Atualizar tudo" : "Escolher pasta"}</button>`;
    }
    $("#ferr").innerHTML = PASTA.erros.map((e) => `<span>${esc(e)}</span>`).join("");
  }

  /* ---------- dados fictícios (um por plano) ---------- */
  function rng(seed) {
    return () => {
      seed |= 0;
      seed = (seed + 0x6d2b79f5) | 0;
      let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  /* Dados fictícios de exemplo (uma conta inventada por plano). */
  function sampleState() {
    const r = rng(1911);
    const rnd = (a, b) => Math.round((a + (b - a) * r()) * 100) / 100;
    const end = addDays(fromD(new Date()), -1);
    const start = end.slice(0, 5) + "01-02";
    const DEF = {
      financeiro: {
        name: "Exemplo BB - 000.001",
        daily: [
          [
            "1215015101 - CONTRIBUIÇÃO SERVIDOR INATIVO – SENTENÇA JUDICIAL (EXEMPLO)",
            0.9,
            6,
            150,
            60000,
            "Precatório",
          ],
        ],
        month: [
          [
            "1999030101 - COMPENSAÇÃO FINANCEIRA RGPS/RPPS (EXEMPLO)",
            6,
            9e6,
            14e6,
            "INSS",
            "COMPREV · União",
          ],
          ["1718011101 - ROYALTIES DO PETRÓLEO (EXEMPLO)", 15, 6e7, 9e7, "Repasse de royalties", "Royalties"],
        ],
        desp: [
          ["319001 - Aposentadorias e Reformas", 26, 1.1e8, 1.3e8, "Folha de inativos – Educação"],
          ["319001 - Aposentadorias e Reformas", 26, 4e7, 5e7, "Folha de inativos – Saúde"],
          ["319003 - Pensões do RPPS e do Militar", 27, 3e7, 3.8e7, "Folha de pensionistas"],
          ["319086 - Compensações a regimes de previdência", 20, 2e6, 4e6, "Órgão federal exemplo"],
        ],
      },
      previdenciario: {
        name: "Exemplo BB - 000.002",
        daily: [],
        month: [
          [
            "1215011101 - CONTRIBUIÇÃO PATRONAL – SERVIDOR ATIVO (EXEMPLO)",
            10,
            2.2e7,
            2.6e7,
            "Tesouro Estadual",
            "Contribuição",
          ],
          [
            "1215012101 - CONTRIBUIÇÃO DO SERVIDOR ATIVO (EXEMPLO)",
            10,
            1.4e7,
            1.7e7,
            "Servidores ativos",
            "Contribuição",
          ],
          [
            "1321001101 - REMUNERAÇÃO DE APLICAÇÕES FINANCEIRAS (EXEMPLO)",
            3,
            3e6,
            6e6,
            "Fundos de investimento",
            "Rendimentos",
          ],
        ],
        desp: [
          ["319001 - Aposentadorias e Reformas", 26, 6e6, 8e6, "Folha de aposentados – Previdenciário"],
          [
            "319003 - Pensões do RPPS e do Militar",
            27,
            1.5e6,
            2.2e6,
            "Folha de pensionistas – Previdenciário",
          ],
        ],
      },
      militar: {
        name: "Exemplo BB - 999.990",
        daily: [],
        month: [
          [
            "1215031101 - CONTRIBUIÇÃO DO MILITAR ATIVO (EXEMPLO)",
            10,
            9e6,
            1.1e7,
            "Militares ativos",
            "Contribuição",
          ],
          [
            "1215032101 - CONTRIBUIÇÃO DO MILITAR INATIVO (EXEMPLO)",
            10,
            4e6,
            5e6,
            "Militares inativos",
            "Contribuição",
          ],
        ],
        desp: [
          ["319001 - Aposentadorias e Reformas", 26, 3.2e7, 3.8e7, "Folha de reserva remunerada e reforma"],
          ["319003 - Pensões do RPPS e do Militar", 27, 1.2e7, 1.5e7, "Folha de pensionistas militares"],
        ],
      },
      taxa: {
        name: "Exemplo BB - 000.004",
        daily: [],
        month: [
          [
            "7215019901 - TAXA DE ADMINISTRAÇÃO (EXEMPLO)",
            8,
            5e6,
            6e6,
            "Planos Financeiro, Militar e Previdenciário",
            "Taxa",
          ],
        ],
        desp: [
          ["339039 - Outros serviços de terceiros – PJ", 15, 8e5, 1.4e6, "Prestador de serviços exemplo"],
          ["339040 - Serviços de tecnologia da informação", 18, 4e5, 7e5, "Empresa de TI exemplo"],
          ["339030 - Material de consumo", 12, 4e4, 1.2e5, "Fornecedor exemplo"],
          ["319011 - Vencimentos – pessoal civil", 28, 1.8e6, 2.1e6, "Folha do Rioprevidência"],
        ],
      },
    };
    const accounts = {};
    let n = 0;
    for (const [plano, df] of Object.entries(DEF)) {
      const txs = [];
      for (let k = start; k <= end; k = addDays(k, 1)) {
        const wd = toD(k).getDay();
        if (wd === 0 || wd === 6) continue;
        const day = +k.slice(8),
          last = lastDayOf(k);
        for (const [nat, prob, maxN, a, b, idk] of df.daily) {
          if (r() > prob) continue;
          const q = Math.floor(r() * maxN) + 1;
          for (let i = 0; i < q; i++)
            txs.push({
              date: k,
              tipo: "RECEITA",
              value: Math.round((a + Math.pow(r(), 3) * (b - a)) * 100) / 100,
              nat,
              ident: `Precatório 20${20 + Math.floor(r() * 6)}.${String(Math.floor(r() * 1e5)).padStart(5, "0")}`,
              idk,
              lanc: "Resgate depósito judicial",
              conc: k >= addDays(end, -2) && r() < 0.5 ? 0 : 1,
              n: String(++n),
            });
        }
        for (const [nat, dd, a, b, ident, idk] of df.month) {
          if (day >= dd && !txs.some((t) => t.nat === nat && t.date.slice(0, 7) === k.slice(0, 7)))
            txs.push({
              date: k,
              tipo: "RECEITA",
              value: rnd(a, b),
              nat,
              ident,
              idk,
              lanc: "Crédito recebido",
              conc: 1,
              n: String(++n),
            });
        }
        for (const [nat, dd, a, b, credor] of df.desp) {
          const tgt = Math.min(dd, +last.slice(8));
          if (
            day >= tgt &&
            !txs.some(
              (t) =>
                t.tipo === "DESPESA" &&
                t.natD === nat &&
                t.credor === credor &&
                t.date.slice(0, 7) === k.slice(0, 7),
            )
          ) {
            txs.push({
              date: k,
              tipo: "DESPESA",
              value: -rnd(a, b),
              nat,
              natD: nat,
              credor,
              item: nat.split(" - ")[1],
              lanc: "Pagamento via OB",
              conc: 1,
              n: String(++n),
            });
          }
        }
      }
      if (plano === "financeiro") {
        const ms = [...new Set(txs.map((t) => t.date.slice(0, 7)))];
        for (const mk of ms) {
          const d1 = txs.find((t) => t.date.startsWith(mk) && +t.date.slice(8) >= 20);
          if (!d1) continue;
          txs.push({
            date: d1.date,
            tipo: "NP",
            value: rnd(9e7, 1.2e8),
            credor: "TESOURO ESTADUAL",
            lanc: "Aporte do Tesouro",
            conc: 1,
            n: String(++n),
          });
          txs.push({
            date: d1.date,
            tipo: "TRANSF",
            value: -rnd(3.4e7, 4.2e7),
            credor: "001 - 0001 - 0009999901",
            lanc: "438 TED",
            conc: 1,
            n: String(++n),
          });
        }
        txs.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
      }
      if (plano === "financeiro")
        for (let i = 0; i < 14; i++) {
          const t = txs[txs.length - 1 - i * 3];
          if (t && t.tipo === "RECEITA" && t.idk === "Precatório") {
            t.nat = "Não Encontrado";
          }
        }
      accounts[df.name] = {
        name: df.name,
        file: "dados fictícios",
        planoHint: plano,
        format: "exemplo",
        txs,
        saldos: [],
      };
    }
    return { accounts, updatedAt: null, embedded: true };
  }

  /* ---------- eventos ---------- */
  window.addEventListener("hashchange", () => route(true));
  document.addEventListener("click", async (e) => {
    const go = e.target.closest("[data-go]");
    if (go) {
      location.hash = go.dataset.go;
      return;
    }
    const pr = e.target.closest("[data-preset]");
    if (pr) {
      ui.preset = pr.dataset.preset;
      ui.limit = 50;
      render();
      return;
    }
    const op = e.target.closest("[data-open]");
    if (op) {
      const k = op.dataset.open;
      ui.open.has(k) ? ui.open.delete(k) : ui.open.add(k);
      const y = window.scrollY;
      render();
      window.scrollTo(0, y);
      return;
    }
    const rp = e.target.closest("[data-rplan]");
    if (rp) {
      ui.rplan = rp.dataset.rplan;
      render();
      return;
    }
    const sp = e.target.closest("[data-sp]");
    if (sp) {
      const a = sp.dataset.sp;
      if (a === "connect") spConnect();
      else if (a === "sync") syncSP(true);
      else if (a === "logout") spLogout();
      else if (a === "edit") {
        SP.edit = true;
        renderSP();
      } else if (a === "cancel") {
        SP.edit = false;
        renderSP();
      } else if (a === "off") {
        if (confirmStep(sp, "Confirmar: desconectar")) spDisconnect();
      } else if (a === "copyuri") {
        try {
          await navigator.clipboard.writeText(spRedirect());
          sp.textContent = "Copiado";
        } catch (err) {
          const c = sp.previousElementSibling;
          const rg = document.createRange();
          rg.selectNodeContents(c);
          const se = getSelection();
          se.removeAllRanges();
          se.addRange(rg);
          sp.textContent = "Selecionado";
        }
        setTimeout(() => {
          sp.textContent = "Copiar";
        }, 2000);
      }
      return;
    }
    const fo = e.target.closest("[data-fold]");
    if (fo) {
      if (fo.dataset.fold === "forget" && !confirmStep(fo, "Confirmar: desconectar")) return;
      ({
        pick: connectFolder,
        allow: allowFolder,
        sync: () => syncFolder(true),
        forget: forgetFolder,
        input: () => $("#foldIn").click(),
      })[fo.dataset.fold]();
      return;
    }
    const x = e.target.closest(".rep .x");
    if (x) {
      x.parentElement.remove();
      return;
    }
    const un = e.target.closest("[data-unnum]");
    if (un) {
      state.aporteNums = (state.aporteNums || []).filter(
        (n) => digits(n).replace(/^0+/, "") !== un.dataset.unnum,
      );
      await save();
      rebuild();
      render();
      return;
    }
    const del = e.target.closest("[data-del]");
    if (del) {
      if (!confirmStep(del, "Confirmar remoção")) return;
      const nm = del.closest("tr").dataset.a;
      delete state.accounts[nm];
      await save();
      rebuild();
      render();
      return;
    }
    const act = e.target.closest("[data-act]");
    if (!act) return;
    const a = act.dataset.act;
    if (a === "load") $("#file").click();
    else if (a === "sample") {
      state = sampleState();
      rebuild();
      ui.open.clear();
      render(true);
      $("#report").innerHTML = "";
    } else if (a === "more") {
      ui.limit += 100;
      renderTx();
    } else if (a === "toap") {
      const c = $("#apCard");
      if (c) {
        const d = c.querySelector("details");
        if (d) d.open = true;
        c.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    } else if (a === "addnum") {
      const v = digits($("#apNum").value);
      if (v.length < 4) {
        $("#apNum").focus();
        return;
      }
      state.aporteNums = [...new Set([...(state.aporteNums || []), v])];
      await save();
      rebuild();
      render();
    } else if (a === "print") window.print();
    else if (a === "copy") {
      const el = $("#doc");
      try {
        await navigator.clipboard.writeText(el.innerText);
        act.textContent = "Copiado";
      } catch (err) {
        const rg = document.createRange();
        rg.selectNodeContents(el);
        const s = getSelection();
        s.removeAllRanges();
        s.addRange(rg);
        act.textContent = "Selecionado: use Ctrl+C";
      }
      setTimeout(() => {
        act.textContent = "Copiar texto";
      }, 2500);
    } else if (a === "clear") {
      if (!confirmStep(act, "Confirmar: limpar tudo")) return;
      const keep = {
        planoSet: state.planoSet || {},
        aporteDest: state.aporteDest || {},
        aporteOrig: state.aporteOrig || {},
        aporteNums: state.aporteNums || [],
      };
      state = { accounts: {}, updatedAt: null, embedded: false, ...keep };
      await save();
      rebuild();
      render();
      $("#report").innerHTML = "";
    }
  });
  document.addEventListener("change", async (e) => {
    if (e.target.matches("#dFrom,#dTo")) {
      ui.preset = "custom";
      ui.from = $("#dFrom").value || null;
      ui.to = $("#dTo").value || null;
      render();
      return;
    }
    if (e.target.matches("#rsub")) {
      ui.rsub = e.target.checked;
      render();
      return;
    }
    if (e.target.matches("[data-orig]")) {
      state.aporteOrig = state.aporteOrig || {};
      state.aporteOrig[e.target.dataset.orig] = e.target.checked;
      const y = window.scrollY;
      await save();
      rebuild();
      render();
      window.scrollTo(0, y);
      return;
    }
    if (e.target.matches("[data-dest]")) {
      state.aporteDest = state.aporteDest || {};
      state.aporteDest[e.target.dataset.dest] = e.target.checked;
      const y = window.scrollY;
      await save();
      rebuild();
      render();
      window.scrollTo(0, y);
      return;
    }
    if (e.target.matches("[data-plano]")) {
      const nm = e.target.closest("tr").dataset.a;
      const acc = state.accounts[nm];
      if (!acc) return;
      state.planoSet = state.planoSet || {};
      if (e.target.value) state.planoSet[acc.srcName || acc.name] = e.target.value;
      else delete state.planoSet[acc.srcName || acc.name];
      await save();
      rebuild();
      render();
      return;
    }
  });
  let qt;
  document.addEventListener("input", (e) => {
    if (e.target.matches("#q")) {
      clearTimeout(qt);
      qt = setTimeout(() => {
        ui.q = e.target.value;
        ui.limit = 50;
        renderTx();
      }, 140);
    }
  });
  document.addEventListener("keydown", (e) => {
    if (e.target.matches("#apNum") && e.key === "Enter") {
      e.preventDefault();
      document.querySelector('[data-act="addnum"]').click();
      return;
    }
    if (e.target.matches(".drop-zone") && (e.key === "Enter" || e.key === " ")) {
      e.preventDefault();
      $("#file").click();
    }
  });
  let armT = null;
  function confirmStep(btn, label) {
    if (btn.dataset.armed) {
      delete btn.dataset.armed;
      btn.classList.remove("danger");
      btn.textContent = btn.dataset.orig;
      clearTimeout(armT);
      return true;
    }
    btn.dataset.orig = btn.textContent;
    btn.dataset.armed = "1";
    btn.classList.add("danger");
    btn.textContent = label;
    armT = setTimeout(() => {
      delete btn.dataset.armed;
      btn.classList.remove("danger");
      btn.textContent = btn.dataset.orig;
    }, 3500);
    return false;
  }
  $("#btnLoad").onclick = () => {
    if (ui.view !== "planilhas") location.hash = "planilhas";
    else $("#file").click();
  };
  $("#btnSync").onclick = updateAll;
  $("#file").onchange = (e) => {
    const l = e.target.files;
    if (l && l.length) handleFiles(l);
    e.target.value = "";
  };
  $("#foldIn").onchange = (e) => {
    const l = e.target.files;
    if (l && l.length) folderFromInput(l);
    e.target.value = "";
  };
  const autoSync = () => {
    if (document.visibilityState !== "visible") return;
    if (SP.cfg.active) syncSP(false);
    else syncFolder(false);
  };
  document.addEventListener("visibilitychange", autoSync);
  setInterval(autoSync, 5 * 60 * 1000);
  let dragN = 0;
  const drop = $("#drop");
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes("Files");
  window.addEventListener("dragenter", (e) => {
    if (hasFiles(e)) {
      e.preventDefault();
      dragN++;
      drop.hidden = false;
    }
  });
  window.addEventListener("dragover", (e) => {
    if (hasFiles(e)) e.preventDefault();
  });
  window.addEventListener("dragleave", () => {
    dragN = Math.max(0, dragN - 1);
    if (!dragN) drop.hidden = true;
  });
  window.addEventListener("drop", (e) => {
    if (!e.dataTransfer?.files?.length) return;
    e.preventDefault();
    dragN = 0;
    drop.hidden = true;
    handleFiles(e.dataTransfer.files);
  });
  let rz;
  window.addEventListener("resize", () => {
    clearTimeout(rz);
    rz = setTimeout(() => {
      movePill();
      if (chartArgs && document.body.contains(chartArgs[0])) drawChart(...chartArgs);
    }, 120);
  });
  if (document.fonts && document.fonts.ready)
    document.fonts.ready.then(() => {
      firstPill = true;
      movePill();
    });

  /* ---------- início ---------- */
  try {
    const p = JSON.parse(localStorage.getItem(STORE + "-prefs2") || "null");
    if (p) {
      if (p.preset && p.preset !== "custom") ui.preset = p.preset;
      if (p.rplan) ui.rplan = p.rplan;
      ui.rsub = !!p.rsub;
    }
  } catch (e) {}
  route(false);
  (async () => {
    const ok = await load();
    if (!ok) state = { accounts: {}, updatedAt: null, embedded: false };
    rebuild();
    if (ok) await save();
    if (PASTA.ok) {
      try {
        const h = await IDB.get("pasta");
        if (h) {
          PASTA.handle = h;
          PASTA.perm = await h.queryPermission({ mode: "read" });
        }
      } catch (e) {}
    }
    booting = false;
    render();
    renderFolder();
    if (SP.cfg.active && SP.ok) syncSP(false);
    else if (PASTA.handle && PASTA.perm === "granted") syncFolder(false);
    const velhas = ORDER.filter((n) => {
      const a = state.accounts[n];
      return a.src !== "pasta" && (a.pv || 1) < PARSER_V;
    });
    const pastaVelha =
      Object.values(state.files || {}).some((f) => !f.ignored && (f.pv || 1) < PARSER_V) &&
      !(PASTA.handle && PASTA.perm === "granted");
    if (!state.embedded && (velhas.length || pastaVelha))
      showReport([
        {
          ok: false,
          file: "Aportes (NP)",
          msg: `o painel agora conta as linhas NP como aportes. ${velhas.length ? `Carregue de novo ${velhas.length === 1 ? "a planilha" : "as planilhas"} ${esc(velhas.join(", "))} ` : ""}${pastaVelha ? (velhas.length ? "e clique em <b>Atualizar tudo</b> " : "Clique em <b>Atualizar tudo</b> ") : ""}para incluir as NP que já estavam nelas.`,
        },
      ]);
  })();
})();
