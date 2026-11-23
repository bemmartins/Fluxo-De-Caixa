/* =====================================================================
   Fluxo de Caixa – Tesouraria · leitor das planilhas
   Lê a aba de conciliação (C.C) ou um extrato simples e devolve os lançamentos
   (data, tipo, valor, natureza, credor, identificação…) e os saldos.
   Também reconhece o plano (Taxa, Financeiro, Militar, Previdenciário).
   Ponto de entrada: parseBook(abas, nomeDoArquivo).
   ===================================================================== */
/* ===== Leitor de extratos (sem DOM) ===== */
const norm = (s) =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
const pad = (n) => String(n).padStart(2, "0");
const key = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;
const fromD = (d) => key(d.getFullYear(), d.getMonth() + 1, d.getDate());
/* Converte texto de valor no formato brasileiro (1.234,56 / (123,45) / 123,45 D) em número. */
function parseNum(x) {
  if (x == null || x === "") return NaN;
  if (typeof x === "number") return x;
  let s = String(x).trim();
  if (!s) return NaN;
  const neg = /^\(.*\)$/.test(s) || /-|−/.test(s) || /[dD]\s*$/.test(s);
  s = s.replace(/[^\d.,]/g, "");
  if (!/\d/.test(s)) return NaN;
  const lc = s.lastIndexOf(","),
    ld = s.lastIndexOf(".");
  if (lc > -1 && ld > -1) {
    s = lc > ld ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  } else if (lc > -1) {
    s = s.split(",").length > 2 ? s.replace(/,/g, "") : s.replace(",", ".");
  } else if (ld > -1 && /^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, "");
  }
  const n = parseFloat(s);
  return isNaN(n) ? NaN : neg ? -n : n;
}
function validKey(y, m, d) {
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1990 || y > 2100) return null;
  return key(y, m, d);
}
/* Converte data do Excel (número de série), "dd/mm/aaaa" ou "aaaa-mm-dd" para a chave "aaaa-mm-dd". */
function parseDate(x) {
  if (x == null || x === "") return null;
  if (x instanceof Date && !isNaN(x)) return fromD(x);
  if (typeof x === "number") {
    if (x > 20000 && x < 80000) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(x) * 864e5);
      return key(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
    }
    return null;
  }
  const s = String(x).trim();
  let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return validKey(+m[1], +m[2], +m[3]);
  if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})\b/))) {
    let y = +m[3];
    if (y < 100) y += 2000;
    return validKey(y, +m[2], +m[1]);
  }
  return null;
}
const clean = (v) => {
  if (v == null) return "";
  const s = String(v).replace(/\s+/g, " ").trim();
  return /^[\s\-–]*$/.test(s) || s === "None" ? "" : s;
};

/* UG a partir do histórico, igual à fórmula da coluna UG da planilha */
/* Descobre a UG pelo histórico do lançamento (mesma regra da fórmula da coluna UG da planilha). */
function ugFromLanc(lanc) {
  const L = String(lanc || ""),
    U = L.toUpperCase();
  let i;
  const num = (s) => {
    const n = parseInt(String(s).trim(), 10);
    return isNaN(n) ? "" : String(n);
  };
  if ((i = U.indexOf("PAGAMENTO GOVERNO")) > -1) {
    const j = U.indexOf("UG");
    return j > -1 ? num(L.substr(j + 2, 6)) : "";
  }
  if ((i = U.indexOf("CONTAS")) > -1 && L.length > 20) return num(L.substr(i + 7, 6));
  if ((i = U.indexOf("TRANSF POR OB")) > -1 && L.length > 13) return num(L.substr(i + 14, 6));
  if (U.includes("EDUC")) return "180100";
  if (U.includes("IPEM")) return "53100";
  if (U.includes("SECRETARIA DE ESTADO")) return "180100";
  return "";
}

/* Identificação da receita a partir do texto da coluna RECEITA */
/* Lê a coluna RECEITA e devolve a identificação: ente do COMPREV, número do precatório ou beneficiário. */
function identOf(txt) {
  if (!txt) return { ident: "", kind: "", comp: "", proc: "" };
  const kv = {};
  for (const raw of String(txt).replace(/"/g, "").split(/\r?\n/)) {
    const p = raw.indexOf(":");
    if (p < 1) continue;
    const k = norm(raw.slice(0, p)).replace(/[.\s]+$/, "");
    const v = raw
      .slice(p + 1)
      .replace(/\s+/g, " ")
      .trim();
    if (k && v && !(k in kv)) kv[k] = v;
  }
  const ente = kv["nome do ente"] || kv["ente"];
  if (ente) {
    const t = kv["tipo do ente"];
    const TE = { municipio: "Município", estado: "Estado", uniao: "União", distrito: "Distrito" };
    return {
      ident: ente.replace(/\s+/g, " ").trim(),
      kind: "COMPREV" + (t ? " · " + (TE[norm(t)] || t) : ""),
      comp: kv["competencia"] || "",
      proc: kv["processo sei"] || "",
    };
  }
  if (kv["precatorio"])
    return {
      ident: "Precatório " + kv["precatorio"],
      kind: "Precatório",
      comp: kv["data do levantamento"] || "",
      proc: kv["processo"] || "",
    };
  if (kv["beneficiario"])
    return { ident: kv["beneficiario"], kind: "Resgate", comp: "", proc: kv["processo"] || "" };
  const first =
    String(txt)
      .replace(/"/g, "")
      .split(/\r?\n/)
      .map((s) => s.trim())
      .find(Boolean) || "";
  return { ident: first.slice(0, 70), kind: "", comp: "", proc: "" };
}

const TIPOS = { receita: "RECEITA", despesa: "DESPESA", fi: "FI" };
/* Normaliza a coluna Tipo: RECEITA, DESPESA, TRANSF, FI, NP (aporte) ou OUTROS. */
function tipoOf(t) {
  const n = norm(t).replace(/\.$/, "").trim();
  if (TIPOS[n]) return TIPOS[n];
  if (/^transf/.test(n)) return "TRANSF";
  if (/^(np|n\.p|aporte)/.test(n)) return "NP";
  return n ? "OUTROS" : "";
}
const PARSER_V = 2;

/* Cabeçalho do formato de conciliação (aba C.C) */
const CC_MAP = [
  ["n", /^n\.?$|^no\.?$/],
  ["tipo", /^tipo$/],
  ["item", /^tipo item patrimonial$/],
  ["nat", /^natureza$/],
  ["doc", /^doc\.?$|^documento$/],
  ["credor", /credor|^entidade/],
  ["pd", /^pd$/],
  ["ug", /^ug$/],
  ["comp", /^comp\.?$|^competencia$/],
  ["date", /^data$/],
  ["lanc", /^lancamento$/],
  ["ob", /^ob$/],
  ["value", /^valor$/],
  ["saldo", /^saldo$/],
  ["receita", /^receita$/],
  ["natD", /^natureza da despesa$/],
  ["conc", /^conciliado$/],
  ["analista", /^analista$/],
];
/* Procura a linha de cabeçalho da aba de conciliação (Tipo, DATA, LANÇAMENTO, VALOR…). */
function findCC(rows) {
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const r = rows[i] || [];
    const col = {};
    r.forEach((c, j) => {
      const h = norm(c).replace(/[:]+$/, "").trim();
      if (!h) return;
      for (const [k, re] of CC_MAP) {
        if (col[k] == null && re.test(h)) {
          col[k] = j;
          break;
        }
      }
    });
    if (col.tipo != null && col.date != null && col.value != null && col.lanc != null) return { i, col };
  }
  return null;
}
function titleOf(rows) {
  for (let i = 0; i < Math.min(rows.length, 8); i++)
    for (const c of rows[i] || []) {
      const m = String(c ?? "").match(/extrato banc[aá]rio\s*[-–]\s*(.+)/i);
      if (m) return m[1].trim();
    }
  return "";
}
/* Nome da planilha a partir do nome do arquivo, sem datas, mês, trimestre ou ano. */
function acctName(fname) {
  const raw = String(fname)
    .replace(/^.*[\/\\]/, "")
    .replace(/\.[^.]+$/, "");
  let s = raw.replace(/[_]+/g, " ");
  s = s
    .replace(/\b\d{4}[-.]\d{1,2}[-.]\d{1,2}\b/g, "")
    .replace(/\b\d{1,2}[-.]\d{1,2}[-.]\d{2,4}\b/g, "")
    .replace(/\b(extrato|planilha|relatorio|relatório)\b/gi, "")
    .replace(/\s*[-–]\s*20\d\d\b.*$/, "")
    .replace(/\b\d\s*[ºo°]?\s*(tri|trim|trimestre|sem|semestre)\b\.?/gi, "")
    .replace(
      /\b(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\b/gi,
      "",
    )
    .replace(/\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[-.\/ ]?(20)?\d\d\b/gi, "")
    .replace(/\b20\d\d\b/g, "")
    .replace(/\(\d+\)/g, "")
    .replace(/(\s*[-–·]\s*){2,}/g, " - ")
    .replace(/(\s*[-–·]\s*)+$/, "")
    .replace(/^(\s*[-–·]\s*)+/, "")
    .replace(/\s{2,}/g, " ")
    .trim();
  return s || raw;
}

/* Lê as linhas da aba de conciliação e devolve os lançamentos e os saldos. */
function parseCC(rows, h) {
  const c = h.col,
    txs = [],
    saldos = [];
  const g = (r, k) => (c[k] != null ? r[c[k]] : null);
  for (let k = h.i + 1; k < rows.length; k++) {
    const r = rows[k] || [];
    const date = parseDate(g(r, "date"));
    if (!date) continue;
    const lanc = clean(g(r, "lanc"));
    const v = parseNum(g(r, "value"));
    const s = parseNum(g(r, "saldo"));
    if (/^saldo/.test(norm(lanc)) || isNaN(v) || v === 0) {
      if (!isNaN(s)) saldos.push({ date, saldo: s });
      continue;
    }
    let tipo = tipoOf(g(r, "tipo"));
    if (!tipo) tipo = v > 0 ? "RECEITA" : "DESPESA";
    let ug = clean(g(r, "ug"));
    if (ug && /^\d+(\.0+)?$/.test(ug)) ug = String(parseInt(ug, 10));
    if (!ug && v > 0) ug = ugFromLanc(lanc);
    const id =
      tipo === "RECEITA" || tipo === "NP"
        ? identOf(g(r, "receita"))
        : { ident: "", kind: "", comp: "", proc: "" };
    const conc = norm(g(r, "conc"));
    const comp = parseDate(g(r, "comp"));
    let credor = clean(g(r, "credor"));
    if (!credor && tipo === "FI")
      credor = (v > 0 ? "Resgate" : "Aplicação") + " - " + lanc.replace(/^\d+\s+/, "");
    txs.push({
      date,
      tipo,
      value: Math.round(v * 100) / 100,
      saldo: isNaN(s) ? null : s,
      item: clean(g(r, "item")),
      nat: clean(g(r, "nat")),
      natD: clean(g(r, "natD")),
      credor,
      ug,
      doc: clean(g(r, "doc")),
      ob: clean(g(r, "ob")),
      pd: clean(g(r, "pd")),
      lanc,
      comp: comp ? comp.slice(0, 7) : "",
      ident: id.ident,
      idk: id.kind,
      idcomp: id.comp,
      proc: id.proc,
      conc: conc === "sim" ? 1 : conc === "nao" ? 0 : null,
      n: clean(g(r, "n")),
    });
    if (!isNaN(s)) saldos.push({ date, saldo: s });
  }
  return { txs, saldos };
}

/* Formato genérico de extrato (Data + Valor ou Crédito/Débito) */
/* Procura cabeçalho de extrato simples (Data + Valor, ou Crédito/Débito). */
function findGeneric(rows) {
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const r = rows[i] || [];
    const col = { desc: [] };
    r.forEach((cell, j) => {
      const c = norm(cell);
      if (!c || c.length > 40) return;
      if (col.date == null && /^(data|dt\b|date|dia\b)/.test(c)) {
        col.date = j;
        return;
      }
      if (col.saldo == null && /saldo/.test(c)) {
        col.saldo = j;
        return;
      }
      if (col.credit == null && /credito|entrada/.test(c)) {
        col.credit = j;
        return;
      }
      if (col.debit == null && /debito|saida/.test(c)) {
        col.debit = j;
        return;
      }
      if (col.dc == null && /^(d\/c|c\/d|dc|cd|natureza|deb\/cred|sinal)$/.test(c)) {
        col.dc = j;
        return;
      }
      if (col.value == null && /valor|montante|quantia|amount|^r\$|^vlr/.test(c)) {
        col.value = j;
        return;
      }
      if (/hist|descr|lancament|complement|memo|detalhe|favorecido|identifica/.test(c) && col.desc.length < 3)
        col.desc.push(j);
    });
    if (col.date != null && (col.value != null || col.credit != null || col.debit != null)) return { i, col };
  }
  return null;
}
/* Lê extrato simples: valor positivo vira receita, negativo vira despesa. */
function parseGeneric(rows, h) {
  const c = h.col,
    txs = [],
    saldos = [];
  for (let k = h.i + 1; k < rows.length; k++) {
    const r = rows[k] || [];
    const date = parseDate(r[c.date]);
    if (!date) continue;
    const desc = c.desc
      .map((j) => clean(r[j]))
      .filter(Boolean)
      .join(" · ");
    let v;
    if (c.value != null) {
      v = parseNum(r[c.value]);
      if (c.dc != null && !isNaN(v)) {
        const t = norm(r[c.dc]);
        if (t.startsWith("d")) v = -Math.abs(v);
        else if (t.startsWith("c")) v = Math.abs(v);
      }
    } else v = Math.abs(parseNum(r[c.credit]) || 0) - Math.abs(parseNum(r[c.debit]) || 0);
    const s = c.saldo != null ? parseNum(r[c.saldo]) : NaN;
    if (/^saldo/.test(norm(desc)) || !v || isNaN(v)) {
      if (!isNaN(s)) saldos.push({ date, saldo: s });
      continue;
    }
    txs.push({
      date,
      tipo: v > 0 ? "RECEITA" : "DESPESA",
      value: Math.round(v * 100) / 100,
      saldo: isNaN(s) ? null : s,
      item: "",
      nat: "",
      natD: "",
      credor: "",
      ug: v > 0 ? ugFromLanc(desc) : "",
      doc: "",
      ob: "",
      pd: "",
      lanc: desc || "(sem histórico)",
      comp: "",
      ident: "",
      idk: "",
      idcomp: "",
      proc: "",
      conc: null,
      n: "",
    });
    if (!isNaN(s)) saldos.push({ date, saldo: s });
  }
  return { txs, saldos };
}

/* ---------- plano (Taxa, Financeiro, Militar, Previdenciário) ---------- */
const PLANO_RX = {
  financeiro: /plano financeiro|fundo em reparti|reparticao simples|repartição simples/i,
  previdenciario: /plano previdenci|fundo em capitaliza|fundo previdenci/i,
  militar:
    /plano militar|protecao social dos militares|proteção social dos militares|\bspsm\b|fundo militar/i,
  taxa: /taxa de administra|reserva administrativa|plano taxa|fundo de taxa/i,
};
/* Reconhece o plano pelo nome do arquivo ou da pasta. */
function planoFromName(s) {
  const n = " " + norm(s).replace(/[_\-.\/\\()]+/g, " ") + " ";
  if (/ (taxa|adm|administrativo|administrativa|tx) /.test(n)) return "taxa";
  if (/ (mil|militar|militares|spsm) /.test(n)) return "militar";
  if (/ (prev|previdenciario|capitalizacao) /.test(n)) return "previdenciario";
  if (/ (fin|financeiro|reparticao) /.test(n)) return "financeiro";
  return "";
}
/* Reconhece o plano pelo conteúdo da planilha (conta contábil, observações). */
function planoFromBook(sheets) {
  const c = {};
  for (const s of sheets) {
    let k = 0;
    for (const r of s.rows) {
      if (++k > 5000) break;
      for (const v of r || []) {
        if (typeof v !== "string" || v.length < 10) continue;
        for (const p in PLANO_RX) if (PLANO_RX[p].test(v)) c[p] = (c[p] || 0) + 1;
      }
    }
  }
  const best = Object.entries(c).sort((a, b) => b[1] - a[1])[0];
  return best ? best[0] : "";
}
/* Abas que podem ser lidas num arquivo, da melhor para a pior */
const AUX =
  /^(apoio|razao|tab.?siafe|comprev|conciliacao|disponivel|planilha\d*|resumo|dinamica|tabela dinamica)$/;
/* Lista as abas que dá para ler num arquivo, da melhor (C.C) para a pior. */
function tabsOf(sheets) {
  const out = [];
  for (const s of sheets) {
    const n = norm(s.name);
    const h = findCC(s.rows);
    if (h) {
      out.push({
        sheet: s.name,
        kind: "cc",
        h,
        score: (n === "c.c" || n === "cc" ? 3e6 : 2e6) + s.rows.length,
      });
      continue;
    }
    const g = findGeneric(s.rows);
    if (g) out.push({ sheet: s.name, kind: "gen", h: g, score: (AUX.test(n) ? 0 : 1e6) + s.rows.length });
  }
  return out.sort((a, b) => b.score - a.score);
}
function titleOfBook(sheets) {
  for (const s of sheets) {
    const t = titleOf(s.rows);
    if (t) return t;
  }
  return "";
}
/* Lê uma aba específica (ou a melhor, se sheetName vier vazio) */
/* Lê uma aba específica (ou a melhor) e devolve a planilha pronta. */
function parseTab(sheets, fileName, sheetName) {
  const tabs = tabsOf(sheets);
  const t = sheetName ? tabs.find((x) => x.sheet === sheetName) : tabs[0];
  if (!t) return null;
  const rows = sheets.find((s) => s.name === t.sheet).rows;
  const r = t.kind === "cc" ? parseCC(rows, t.h) : parseGeneric(rows, t.h);
  return {
    name: acctName(fileName),
    title: titleOfBook(sheets),
    planoHint: planoFromName(fileName) || planoFromBook(sheets),
    format: t.kind === "cc" ? "conciliação" : "extrato",
    sheet: t.sheet,
    tabs: tabs.map((x) => x.sheet),
    ...r,
  };
}
/* sheets: [{name, rows}] → a conta do arquivo (uma por arquivo) */
/* Ponto de entrada do leitor: recebe as abas do arquivo e devolve a planilha lida. */
function parseBook(sheets, fileName) {
  for (const t of tabsOf(sheets)) {
    const a = parseTab(sheets, fileName, t.sheet);
    if (a && a.txs.length) return [a];
  }
  return [];
}
