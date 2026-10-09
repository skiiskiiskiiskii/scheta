// Страница: экраны (ключ → выбор таблицы → сопоставление → работа), предпросмотр, правки, запись, настройки.
import { analyze, applyEdits, askJson, validateKey, keyProblem, worst, MODELS, DEFAULT_MODEL, WARN, EDITED } from "./core.js";
import { kv, files, cache, askPersistent, wipeAll, emptyTable } from "./store.js";
import { exportOurTable, readOurTable, downloadBlob, rowNotes } from "./excel.js";
import * as own from "./owntable.js";

const $ = (id) => document.getElementById(id);
const S = { settings: {}, mode: null, table: emptyTable(), own: null, results: [], include: {}, edits: {},
  queue: [], done: 0, total: 0, busy: false, map: null };

// ───────────────────────────── мелочи ─────────────────────────────
function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined || v === false) continue;
    if (k === "class") el.className = v;
    else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const kid of kids.flat()) if (kid !== null && kid !== undefined && kid !== false)
    el.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
  return el;
}
const nf = (d) => new Intl.NumberFormat("ru-RU", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtNum = (v, field) => (v === null || v === undefined ? "" : field === "Литры" && Number.isInteger(v) ? nf(0).format(v) : nf(2).format(v));
const isoToRu = (v) => (v ? v.split("-").reverse().join(".") : "");
function ruToIso(s) {
  s = String(s).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})$/);
  if (!m) return NaN;
  const d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  if (d.getUTCDate() !== +m[1] || d.getUTCMonth() !== +m[2] - 1) return NaN;
  return d.toISOString().slice(0, 10);
}
function parseNum(s) {
  s = String(s).replace(/[\s  ']/g, "");
  if (!s) return null;
  if (s.includes(",") && s.includes(".")) {
    const dec = s.lastIndexOf(",") > s.lastIndexOf(".") ? "," : ".";
    s = s.split(dec === "," ? "." : ",").join("").replace(dec, ".");
  } else s = s.replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}
function plural(n, one, few, many) {
  const m10 = n % 10, m100 = n % 100;
  return m10 === 1 && m100 !== 11 ? one : m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14) ? few : many;
}
const pad = (n) => String(n).padStart(2, "0");
const todayIso = () => { const t = new Date(); return `${t.getFullYear()}-${pad(t.getMonth() + 1)}-${pad(t.getDate())}`; };
const stamp = () => { const t = new Date(); return `${pad(t.getDate())}.${pad(t.getMonth() + 1)}.${t.getFullYear()} ${pad(t.getHours())}:${pad(t.getMinutes())}`; };
const STATUS_LABEL = { "OK": "OK", "ПРОВЕРИТЬ": "Проверить", "ОШИБКА": "Ошибка", "ИСПРАВЛЕНО": "Исправлено" };
const RANK = { "OK": 0, "ИСПРАВЛЕНО": 1, "ПРОВЕРИТЬ": 2, "ОШИБКА": 3 };
const STRIPE = { "OK": "", "ИСПРАВЛЕНО": "s-edit", "ПРОВЕРИТЬ": "s-warn", "ОШИБКА": "s-err" };

function show(view) {
  for (const v of ["viewKey", "viewChoose", "viewMap", "viewWork"]) $(v).hidden = v !== view;
  $("headActions").hidden = view !== "viewWork";
  updateDock();
}

// ───────────────────────────── сохранение ─────────────────────────────
const saveTable = () => kv.set("table", S.table);
const saveOwn = () => kv.set("own", S.own);
let pendingTimer = null;
function savePending() {
  clearTimeout(pendingTimer);
  pendingTimer = setTimeout(() => kv.set("pending", { results: S.results, include: S.include, edits: S.edits }).catch(() => {}), 250);
}

function unsavedRows() {
  if (S.mode === "own" && S.own) return S.own.appended - S.own.exportedAppended;
  return S.table.rows.length - (S.table.exportedRows || 0);
}

function renderChrome() {
  const chip = $("tableChip");
  if (!S.mode) { chip.replaceChildren(); return; }
  chip.replaceChildren(S.mode === "own"
    ? h("span", {}, "Таблица: ", h("b", {}, S.own.name), ` · добавлено с сайта: ${S.own.appended}`)
    : h("span", {}, "Таблица: ", h("b", {}, "Счета.xlsx"), ` · строк: ${S.table.rows.length}`));
  const n = unsavedRows();
  $("remind").hidden = n <= 0;
  $("remindText").textContent = `С последнего скачивания добавлено ${n} ${plural(n, "строка", "строки", "строк")}. Скачайте таблицу, чтобы файл на компьютере был актуальным.`;
}

// ───────────────────────────── 1. ключ ─────────────────────────────
async function saveKey(input, errEl, after) {
  const key = input.value.trim();
  errEl.textContent = "";
  errEl.className = "err";
  const p = keyProblem(key);
  if (p) { errEl.textContent = p; return; }
  errEl.textContent = "Проверяю ключ…";
  errEl.className = "muted";
  try {
    await validateKey(key);
  } catch (e) {
    errEl.className = "err";
    errEl.textContent = e.message;
    return;
  }
  S.settings.key = key;
  S.settings.model ||= DEFAULT_MODEL;
  await kv.set("settings", S.settings);
  askPersistent();
  input.value = "";
  errEl.className = "okmsg";
  errEl.textContent = "Ключ работает и сохранён.";
  after?.();
}
$("keySave").addEventListener("click", () => saveKey($("keyInput"), $("keyErr"), route));
$("keyInput").addEventListener("keydown", (e) => { if (e.key === "Enter") $("keySave").click(); });

// ───────────────────────────── 2. выбор таблицы ─────────────────────────────
function showChoose(fromSettings = false) {
  $("chooseErr").textContent = "";
  $("chooseBack").hidden = !fromSettings;
  const n = unsavedRows();
  if (fromSettings && n > 0) $("chooseErr").textContent = `Внимание: в текущей таблице ${n} ${plural(n, "строка", "строки", "строк")} ещё не скачаны. Скачайте её перед сменой.`;
  show("viewChoose");
}
$("chooseCancel").addEventListener("click", () => route());
$("chooseNew").addEventListener("click", async () => {
  S.table = emptyTable(); S.mode = "new"; S.own = null;
  await Promise.all([saveTable(), kv.set("mode", "new"), kv.del("own"), files.del("own-workbook")]);
  route();
});
$("ownPicker").addEventListener("change", async (ev) => {
  const f = ev.target.files[0];
  ev.target.value = "";
  if (!f) return;
  $("chooseErr").textContent = "Читаю файл…";
  try {
    const buf = await f.arrayBuffer();
    let ours = null;
    try { ours = await readOurTable(buf.slice(0)); } catch { ours = null; }
    if (ours) {
      S.table = { ...emptyTable(), ...ours }; S.mode = "new"; S.own = null;
      await Promise.all([saveTable(), kv.set("mode", "new"), kv.del("own"), files.del("own-workbook")]);
      route();
      showNotice("good", `Таблица восстановлена: ${ours.rows.length} ${plural(ours.rows.length, "строка", "строки", "строк")}. Дубли из неё тоже будут отсекаться.`);
      return;
    }
    const book = await own.openWorkbook(buf.slice(0));
    const visible = book.sheets.filter((s) => !s.hidden);
    if (!visible.length) throw new Error("в файле нет листов");
    S.map = { name: f.name, buf, book, sheets: visible, sheetIdx: 0 };
    $("chooseErr").textContent = "";
    await openMapping();
  } catch (e) {
    $("chooseErr").textContent = "Не получилось прочитать файл: " + e.message;
  }
});

// ───────────────────────────── 3. сопоставление колонок ─────────────────────────────
async function openMapping() {
  const m = S.map;
  // лист по умолчанию — тот, где больше всего узнаваемых заголовков
  if (m.sheet === undefined) {
    let best = 0, bestScore = -1;
    for (let i = 0; i < m.sheets.length; i++) {
      const sh = await own.readSheet(m.book, m.sheets[i].path);
      const hr = own.detectHeader(sh);
      const score = own.describeColumns(sh, hr).filter((c) => c.field).length;
      if (score > bestScore) { best = i; bestScore = score; }
    }
    m.sheetIdx = best;
  }
  m.sheet = await own.readSheet(m.book, m.sheets[m.sheetIdx].path);
  m.headerRow ??= own.detectHeader(m.sheet);
  m.cols = own.describeColumns(m.sheet, m.headerRow);
  m.ai = {};
  renderMapping();
  show("viewMap");
  suggestWithAI();
}

function renderMapping() {
  const m = S.map;
  $("mapIntro").textContent = `Файл «${m.name}». Сайт сопоставил колонки по названиям — проверьте и поправьте, если нужно. ` +
    "Колонки с «— не заполнять —» останутся пустыми.";
  $("sheetField").hidden = false;
  $("sheetSelect").replaceChildren(...m.sheets.map((s, i) => h("option", { value: i, selected: i === m.sheetIdx }, s.name)));
  const hdrOpts = [];
  for (let r = 1; r <= Math.min(Math.max(m.sheet.maxRow, 1), 30); r++) {
    const cells = m.sheet.rows.get(r);
    const preview = cells ? [...cells.values()].map((c) => own.displayValue(c)).filter(Boolean).slice(0, 4).join(" | ") : "";
    hdrOpts.push(h("option", { value: r, selected: r === m.headerRow }, `${r}: ${preview.slice(0, 60) || "(пусто)"}`));
  }
  $("headerSelect").replaceChildren(...hdrOpts);

  const last = own.lastDataRow(m.sheet, m.headerRow);
  const totals = own.totalsRowHint(m.sheet, m.headerRow);
  $("mapHint").hidden = !totals;
  if (totals) $("mapHint").textContent = `Похоже, строка ${totals} — это строка итогов. Новые строки будут дописываться после неё. ` +
    "Если итоги должны оставаться внизу, напишите — доработаем.";
  $("mapNote").textContent = m.cols.length ? `Новые строки будут дописываться начиная со строки ${last + 1}.` : "";

  const options = (sel) => [h("option", { value: "" }, "— не заполнять —"),
    ...own.FIELDS.map(([k, label]) => h("option", { value: k, selected: k === sel }, label))];
  const rows = m.cols.map((c, i) => h("tr", { class: c.field ? "mapped" : "" },
    h("td", { class: "col" }, c.col),
    h("td", {}, h("b", {}, c.header)),
    h("td", { class: "samples" }, c.samples.join(" · ") || "—"),
    h("td", {}, h("select", { "aria-label": `Что записывать в колонку ${c.header}`, onchange: (ev) => {
      const v = ev.target.value || null;
      if (v) m.cols.forEach((o, j) => { if (j !== i && o.field === v) o.field = null; });
      c.field = v; c.byAI = false;
      renderMapping();
    } }, options(c.field)), c.byAI ? h("span", { class: "muted" }, " подсказка ИИ") : null)));
  $("mapTable").replaceChildren(
    h("thead", {}, h("tr", {}, h("th", {}, ""), h("th", {}, "Заголовок у вас"), h("th", {}, "Последние значения"), h("th", {}, "Что сайт будет туда записывать"))),
    h("tbody", {}, rows));
  if (!m.cols.length) $("mapErr").textContent = "В этой строке нет заголовков — выберите другую строку с заголовками.";
  else $("mapErr").textContent = "";
}

$("sheetSelect").addEventListener("change", async (e) => { S.map.sheetIdx = +e.target.value; S.map.headerRow = undefined; await openMapping(); });
$("headerSelect").addEventListener("change", async (e) => { S.map.headerRow = +e.target.value; await openMapping(); });
$("mapCancel").addEventListener("click", () => { S.map = null; showChoose(!!S.mode); });

async function suggestWithAI() {
  const m = S.map;
  const unknown = m.cols.filter((c) => !c.field);
  if (!unknown.length || !S.settings.key) return;
  const token = (m.aiToken = Math.random());
  $("mapNote").textContent += " ИИ подбирает остальные колонки…";
  const fields = own.FIELDS.map(([k]) => k);
  try {
    const { json } = await askJson(S.settings.key, S.settings.model || DEFAULT_MODEL,
      "You map columns of a user's spreadsheet of supplier invoices (fuel, oils) to known fields. " +
      "Use only the listed field ids, or \"none\" if nothing fits or you are unsure. Each field may be used once. " +
      "Field ids: " + own.FIELDS.map(([k, l]) => `${k} = ${l}`).join("; "),
      JSON.stringify({ already_mapped: m.cols.filter((c) => c.field).map((c) => ({ column: c.col, header: c.header, field: c.field })),
        to_map: unknown.map((c) => ({ column: c.col, header: c.header, examples: c.samples })) }),
      { type: "object", additionalProperties: false, required: ["columns"], properties: { columns: { type: "array", items: {
        type: "object", additionalProperties: false, required: ["column", "field"],
        properties: { column: { type: "string" }, field: { type: "string", enum: [...fields, "none"] } } } } } }, "mapping");
    if (S.map !== m || m.aiToken !== token) return;
    const used = new Set(m.cols.map((c) => c.field).filter(Boolean));
    for (const s of json.columns) {
      const c = m.cols.find((x) => x.col === s.column);
      if (c && !c.field && s.field !== "none" && !used.has(s.field)) { c.field = s.field; c.byAI = true; used.add(s.field); }
    }
    renderMapping();
  } catch {
    if (S.map === m) renderMapping();
  }
}

$("mapSave").addEventListener("click", async () => {
  const m = S.map;
  const mapping = Object.fromEntries(m.cols.filter((c) => c.field).map((c) => [c.col, c.field]));
  if (!Object.keys(mapping).length) { $("mapErr").textContent = "Выберите хотя бы одну колонку, которую заполнять."; return; }
  const sheet = m.sheets[m.sheetIdx];
  S.own = { name: m.name, sheetName: sheet.name, sheetPath: sheet.path, headerRow: m.headerRow, mapping,
    index: own.existingIndex(m.sheet, m.headerRow, mapping), appended: 0, exportedAppended: 0 };
  S.mode = "own";
  S.table = emptyTable();
  await files.set("own-workbook", new Blob([m.buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  await Promise.all([saveOwn(), saveTable(), kv.set("mode", "own")]);
  S.map = null;
  askPersistent();
  route();
});

// ───────────────────────────── 4. работа: загрузка счетов ─────────────────────────────
const drop = $("drop");
["dragenter", "dragover"].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.add("over"); $("dropBig").textContent = "Отпустите, чтобы распознать"; }));
["dragleave", "drop"].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.remove("over"); setDropText(); }));
drop.addEventListener("drop", (ev) => enqueue([...ev.dataTransfer.files]));
$("picker").addEventListener("change", (ev) => { enqueue([...ev.target.files]); ev.target.value = ""; });
document.addEventListener("dragover", (ev) => ev.preventDefault());
document.addEventListener("drop", (ev) => { if (!drop.contains(ev.target)) { ev.preventDefault(); if (!$("viewWork").hidden) enqueue([...ev.dataTransfer.files]); } });

const setDropText = () => { $("dropBig").textContent = S.results.length ? "Добавить ещё счета — перетащите или нажмите" : "Перетащите PDF‑счета сюда"; };

function enqueue(list) {
  const pdfs = list.filter((f) => f.name.toLowerCase().endsWith(".pdf"));
  const other = list.length - pdfs.length;
  if (other) showNotice("bad", `Пропущено файлов не в PDF: ${other}. Сайт принимает только PDF.`); else hideNotice();
  if (!pdfs.length) return;
  if (!S.busy) { S.done = 0; S.total = 0; }
  S.queue.push(...pdfs);
  S.total += pdfs.length;
  if (!S.busy) runQueue();
}

function knownMaps() {
  const keys = new Map(Object.entries(S.table.keys || {})), hashes = new Map(Object.entries(S.table.hashes || {}));
  for (const r of S.results) {
    if (r.sha && !hashes.has(r.sha)) hashes.set(r.sha, `есть в этой пачке (${r.filename})`);
    if (r.kind === "invoice" && r.key && !keys.has(r.key)) keys.set(r.key, `есть в этой пачке (${r.filename})`);
  }
  return { keys, hashes };
}

async function runQueue() {
  S.busy = true;
  $("progress").classList.add("on");
  updateDock();
  while (S.queue.length) {
    const f = S.queue.shift();
    $("progressText").textContent = `Распознаю ${S.done + 1} из ${S.total}: ${f.name}`;
    $("barFill").style.width = `${(S.done / S.total) * 100}%`;
    let res;
    try {
      const bytes = new Uint8Array(await f.arrayBuffer());
      const { keys, hashes } = knownMaps();
      res = await analyze(bytes, f.name, { key: S.settings.key, model: S.settings.model || DEFAULT_MODEL, keys, hashes, today: todayIso(),
        cacheGet: (sha) => cache.get(sha), cachePut: (sha, inv) => cache.set(sha, inv) });
      if (res.kind === "invoice") {
        await files.set("pdf:" + res.id, new Blob([bytes], { type: "application/pdf" }));
        if (S.mode === "own" && S.own?.index) {
          for (const row of res.rows) {
            for (const msg of own.similarRows(S.own.index, res.invoice.number, row)) row.messages.push([WARN, msg]);
            row.status = worst([...res.head, ...row.messages].map(([l]) => l));
          }
          res.status = worst(res.rows.map((r) => r.status));
        }
      }
    } catch (e) {
      res = { id: "x" + Math.random().toString(36).slice(2), filename: f.name, kind: "error", message: "не удалось обработать файл: " + e.message, rows: [], head: [] };
    }
    addResult(res);
    S.done++;
  }
  $("barFill").style.width = "100%";
  const anyNew = S.results.some((r) => r.kind === "invoice");
  $("progressText").textContent = `Готово: распознано ${S.total} ${plural(S.total, "файл", "файла", "файлов")}. ` +
    (anyNew ? "Проверьте строки и нажмите «Добавить в таблицу»." : "Новых счетов среди них нет.");
  S.busy = false;
  updateDock();
}

// ───────────────────────────── предпросмотр ─────────────────────────────
function addResult(res, restoring = false) {
  S.results.push(res);
  if (!(res.id in S.include)) S.include[res.id] = res.include;
  $("emptyHint").hidden = true;
  drop.classList.add("compact");
  setDropText();
  $("results").append(res.kind === "invoice" ? renderInvoice(res) : renderSkip(res));
  if (!restoring) savePending();
  updateDock();
}

function renderSkip(res) {
  const label = { dup_file: "Дубль", dup_invoice: "Дубль", not_invoice: "Не счёт", error: "Сбой" }[res.kind];
  const tail = res.kind === "error" ? "Файл не добавлен — загрузите его ещё раз." : "В таблицу не попадёт.";
  return h("section", { class: "skip" + (res.kind === "error" ? " fail" : ""), id: "inv-" + res.id },
    h("span", { class: "pill " + (res.kind === "error" ? "p-ОШИБКА" : "p-grey") }, label),
    h("span", { class: "fn" }, res.filename),
    h("span", { class: "why" }, `${res.message}. ${tail}`));
}

const editsOf = (id) => (S.edits[id] ||= { invoice: {}, rows: {} });

function cellInput(res, scope, idx, field, original, opts = {}) {
  const isNum = ["Сумма", "Литры", "Цена за л"].includes(field);
  const isDate = field === "Дата" || field === "Дата отгрузки";
  const saved = scope === "invoice" ? S.edits[res.id]?.invoice?.[field] : S.edits[res.id]?.rows?.[idx]?.[field];
  const base = isNum ? fmtNum(original, field) : isDate ? isoToRu(original) : (original ?? "");
  const shown = saved !== undefined ? (isDate ? isoToRu(saved) : saved) : base;
  const inp = h("input", { class: "cell" + (isNum ? " num" : "") + (opts.cls ? " " + opts.cls : "") + (saved !== undefined ? " edited" : ""),
    type: "text", value: shown, inputmode: isNum ? "decimal" : isDate ? "numeric" : null,
    placeholder: isDate ? "дд.мм.гггг" : null, title: shown || null, "aria-label": field + (opts.label ? ": " + opts.label : "") });
  const store = () => {
    const e = editsOf(res.id);
    const bucket = scope === "invoice" ? e.invoice : (e.rows[idx] ||= {});
    let changed;
    if (isNum) {
      const n = parseNum(inp.value);
      inp.classList.toggle("invalid", Number.isNaN(n));
      if (Number.isNaN(n)) return;
      changed = !(n === null && (original ?? null) === null) && (n === null || original === null || original === undefined || Math.abs(n - original) > 1e-9);
      if (changed) bucket[field] = inp.value; else delete bucket[field];
    } else if (isDate) {
      const iso = ruToIso(inp.value);
      inp.classList.toggle("invalid", Number.isNaN(iso));
      if (Number.isNaN(iso)) return;
      changed = iso !== (original || null);
      if (changed) bucket[field] = iso ?? ""; else delete bucket[field];
    } else {
      changed = (inp.value || "") !== (original || "");
      if (changed) bucket[field] = inp.value; else delete bucket[field];
    }
    inp.classList.toggle("edited", changed);
    savePending();
    refreshStatuses(res);
  };
  inp.addEventListener("change", store);
  inp.addEventListener("blur", () => { if (isNum && !inp.classList.contains("invalid") && inp.value) { const n = parseNum(inp.value); if (n !== null) inp.value = fmtNum(n, field); } });
  inp.addEventListener("keydown", (ev) => { if (ev.key === "Enter") inp.blur(); });
  return inp;
}

async function openPdf(res) {
  const blob = await files.get("pdf:" + res.id);
  if (!blob) { showNotice("bad", "PDF этого счёта уже не хранится — откройте его у себя на компьютере."); return; }
  window.open(URL.createObjectURL(blob), "_blank", "noopener");
}

function renderInvoice(res) {
  const inv = res.invoice, first = res.rows[0]?.values || {};
  const cb = h("input", { type: "checkbox", checked: S.include[res.id], onchange: (ev) => {
    S.include[res.id] = ev.target.checked; sec.classList.toggle("off", !ev.target.checked); savePending(); updateDock(); } });
  const pill = h("span", { class: "pill", id: "pill-" + res.id });
  const meta = h("div", { class: "meta" },
    h("label", {}, "Поставщик", cellInput(res, "invoice", null, "Поставщик", first["Поставщик"], { cls: "w-wide" })),
    h("label", {}, "№ счёта", cellInput(res, "invoice", null, "№ счёта", first["№ счёта"], { cls: "w-mid" })),
    h("label", {}, "Дата", cellInput(res, "invoice", null, "Дата", first["Дата"], { cls: "w-date" })),
    h("label", {}, "Дата отгрузки" + (inv.shipment_label ? ` (${inv.shipment_label})` : ""), cellInput(res, "invoice", null, "Дата отгрузки", first["Дата отгрузки"], { cls: "w-date" })),
    h("label", {}, "Валюта", cellInput(res, "invoice", null, "Валюта", first["Валюта"], { cls: "w-s" })));
  const head = res.head.length ? h("div", { class: "head-msgs" }, res.head.map(([lvl, m]) => h("div", { class: "lvl-" + lvl }, m))) : null;
  const tbody = h("tbody", {}, res.rows.map((row, i) => h("tr", { id: `row-${res.id}-${i}`, class: "r-" + row.status },
    h("td", {}, cellInput(res, "rows", i, "Название", row.values["Название"], { label: "строка " + (i + 1) })),
    h("td", { class: "num" }, cellInput(res, "rows", i, "Сумма", row.values["Сумма"])),
    h("td", { class: "num" }, cellInput(res, "rows", i, "Литры", row.values["Литры"])),
    h("td", { class: "num" }, cellInput(res, "rows", i, "Цена за л", row.values["Цена за л"])),
    h("td", {}, h("span", { class: "pill p-" + row.status, id: `rp-${res.id}-${i}` }, STATUS_LABEL[row.status])),
    h("td", { class: "msgs" }, row.messages.map(([lvl, m]) => h("div", { class: "lvl-" + lvl }, m))))));
  const sec = h("section", { class: "invoice" + (S.include[res.id] ? "" : " off"), id: "inv-" + res.id },
    h("div", { class: "inv-top" },
      h("label", {}, cb, res.filename),
      h("a", { href: "#", onclick: (ev) => { ev.preventDefault(); openPdf(res); } }, "Открыть PDF"),
      h("span", { class: "spacer" }), pill),
    meta, head,
    h("div", { class: "table-wrap" }, h("table", {},
      h("colgroup", {}, ["c-name", "c-sum", "c-l", "c-p", "c-st", "c-msg"].map((c) => h("col", { class: c }))),
      h("thead", {}, h("tr", {}, h("th", {}, "Название"), h("th", { class: "num" }, "Сумма" + (first["Валюта"] ? ", " + first["Валюта"] : "")),
        h("th", { class: "num" }, "Литры"), h("th", { class: "num" }, "Цена за л"), h("th", {}, "Статус"), h("th", {}, "Что проверить"))),
      tbody)));
  refreshStatuses(res, sec);
  return sec;
}

function rowStatus(res, i) {
  const e = S.edits[res.id];
  const edited = e && (Object.keys(e.invoice || {}).length || Object.keys(e.rows?.[i] || {}).length);
  return edited ? EDITED : res.rows[i].status;
}

function refreshStatuses(res, root = document) {
  let w = "OK";
  res.rows.forEach((_, i) => {
    const s = rowStatus(res, i);
    if (RANK[s] > RANK[w]) w = s;
    const tr = root.querySelector(`#row-${CSS.escape(res.id)}-${i}`), p = root.querySelector(`#rp-${CSS.escape(res.id)}-${i}`);
    if (tr) tr.className = "r-" + s;
    if (p) { p.className = "pill p-" + s; p.textContent = STATUS_LABEL[s]; }
  });
  const sec = root.querySelector?.("#inv-" + CSS.escape(res.id)) || root;
  if (sec.classList) { sec.classList.remove("s-warn", "s-err", "s-edit"); if (STRIPE[w]) sec.classList.add(STRIPE[w]); }
  const pill = root.querySelector("#pill-" + CSS.escape(res.id));
  if (pill) { pill.className = "pill p-" + w; pill.textContent = STATUS_LABEL[w]; }
  if (root === document) updateDock();
}

function updateDock() {
  const inv = S.results.filter((r) => r.kind === "invoice");
  const chosen = inv.filter((r) => S.include[r.id]);
  const rows = chosen.reduce((s, r) => s + r.rows.length, 0);
  const counts = { "OK": 0, "ПРОВЕРИТЬ": 0, "ОШИБКА": 0, "ИСПРАВЛЕНО": 0 };
  chosen.forEach((r) => r.rows.forEach((_, i) => counts[rowStatus(r, i)]++));
  const skipped = S.results.length - inv.length;
  $("dock").classList.toggle("on", S.results.length > 0 && !$("viewWork").hidden);
  $("dockWhat").textContent = rows
    ? `Будет добавлено ${rows} ${plural(rows, "строка", "строки", "строк")} из ${chosen.length} ${plural(chosen.length, "счёта", "счетов", "счетов")}`
    : inv.length ? "Ни один счёт не отмечен — в таблицу ничего не попадёт" : "Новых счетов нет — в таблицу ничего не попадёт";
  $("commitBtn").textContent = rows ? "Добавить в таблицу" : "Записать в журнал и очистить";
  const parts = [];
  if (counts["ОШИБКА"]) parts.push(`с ошибкой: ${counts["ОШИБКА"]}`);
  if (counts["ПРОВЕРИТЬ"]) parts.push(`проверить: ${counts["ПРОВЕРИТЬ"]}`);
  if (counts["ИСПРАВЛЕНО"]) parts.push(`исправлено вручную: ${counts["ИСПРАВЛЕНО"]}`);
  if (skipped) parts.push(`не попадут (дубли, сбои): ${skipped}`);
  $("dockSub").textContent = parts.join("; ");
  $("commitBtn").disabled = S.busy || !S.results.length;
  const sum = $("summary");
  sum.hidden = !S.results.length;
  sum.replaceChildren(...[
    inv.length ? h("span", {}, h("b", {}, `${inv.length}`), ` ${plural(inv.length, "счёт", "счёта", "счетов")} распознано`) : null,
    ...[["OK", "var(--ok-fg)", "без замечаний"], ["ПРОВЕРИТЬ", "#d6a200", "проверить"], ["ОШИБКА", "var(--err-fg)", "с ошибкой"], ["ИСПРАВЛЕНО", "var(--edit-fg)", "исправлено"]]
      .filter(([k]) => counts[k]).map(([k, c, t]) => h("span", {}, h("i", { class: "sq", style: `background:${c}` }), `${t}: ${counts[k]}`)),
    skipped ? h("span", {}, `не попадут в таблицу: ${skipped}`) : null].filter(Boolean));
}

function showNotice(kind, text, extra) { const n = $("notice"); n.className = "notice on " + kind; n.replaceChildren(...[text, extra].filter(Boolean)); }
function hideNotice() { $("notice").className = "notice"; }

// ───────────────────────────── запись в таблицу ─────────────────────────────
function usageCols(u = {}) {
  return { model: u.model || "", tokens: u.input_tokens !== undefined ? `${u.input_tokens}/${u.output_tokens}` : u.cached ? "из кэша" : "", cost: u.cost_usd || 0 };
}

$("commitBtn").addEventListener("click", async () => {
  const btn = $("commitBtn");
  btn.disabled = true;
  btn.textContent = "Добавляю…";
  try {
    const table = structuredClone(S.table);
    const now = stamp();
    const items = [], starts = [];
    let added = 0, dups = 0, skipped = 0;
    for (const r of S.results) {
      const uc = usageCols(r.usage);
      if (r.kind === "invoice" && S.include[r.id]) {
        const e = applyEdits(r, S.edits[r.id] || {});
        const first = table.rows.length + 2;
        starts.push({ r, idx: items.length, first });
        for (const row of e.rows) {
          const full = { ...row, messages: [...e.head, ...row.messages], file: r.filename, key: r.key, sha: r.sha, added: now };
          table.rows.push(full);
          items.push({ row: full, file: r.filename, notes: rowNotes(full) });
        }
        added += e.rows.length;
        const edited = e.rows.filter((x) => x.edits.length).length;
        table.journal.push({ time: now, file: r.filename, rows: e.rows.length, ...uc,
          result: `добавлено, статус: ${worst(e.rows.map((x) => x.status))}` + (edited ? `, исправлено вручную строк: ${edited}` : "") });
      } else if (r.kind === "invoice") {
        skipped++;
        table.journal.push({ time: now, file: r.filename, rows: 0, ...uc, result: "не добавлен — отменён вручную" });
      } else if (r.kind === "dup_file" || r.kind === "dup_invoice") {
        dups++;
        table.journal.push({ time: now, file: r.filename, rows: 0, ...uc, result: `дубль, не добавлен: ${r.message}` });
      } else {
        table.journal.push({ time: now, file: r.filename, rows: 0, ...uc, result: r.kind === "not_invoice" ? "не похоже на счёт — не добавлен" : r.message });
      }
    }
    let ownState = S.own ? structuredClone(S.own) : null;
    if (S.mode === "own" && items.length) {
      const blob = await files.get("own-workbook");
      const book = await own.openWorkbook(await blob.arrayBuffer());
      const pos = await own.appendRows(book, ownState.sheetPath, ownState.headerRow, ownState.mapping, items);
      await files.set("own-workbook", await own.toBlob(book));
      ownState.appended += items.length;
      for (const s of starts) {
        const where = `в вашей таблице (строка ${pos.first + s.idx}, ${s.r.filename})`;
        if (s.r.key) table.keys[s.r.key] = where;
        table.hashes[s.r.sha] = where;
        const n = s.r.invoice.number ? String(s.r.invoice.number).toLowerCase().replace(/[\W_]/gu, "") : null;
        if (n) ownState.index.numbers[n] ??= pos.first + s.idx;
      }
    } else {
      for (const s of starts) {
        const where = `в таблице (строка ${s.first}, ${s.r.filename})`;
        if (s.r.key) table.keys[s.r.key] = where;
        table.hashes[s.r.sha] = where;
      }
    }
    S.table = table;
    S.own = ownState;
    await Promise.all([saveTable(), S.own ? saveOwn() : null]);
    for (const r of S.results) files.del("pdf:" + r.id).catch(() => {});
    resetList();
    await kv.del("pending");
    const parts = [`Добавлено ${added} ${plural(added, "строка", "строки", "строк")} в ${S.mode === "own" ? "«" + S.own.name + "»" : "таблицу"}`];
    if (dups) parts.push(`дублей пропущено: ${dups}`);
    if (skipped) parts.push(`не добавлено по вашему выбору: ${skipped}`);
    showNotice("good", parts.join("; ") + ".", added ? h("button", { class: "btn", type: "button", onclick: download }, "Скачать таблицу") : null);
    renderChrome();
  } catch (e) {
    showNotice("bad", "Не получилось записать: " + e.message);
  } finally {
    updateDock();
  }
});

function resetList() {
  S.results = []; S.include = {}; S.edits = {};
  $("results").replaceChildren(); $("progress").classList.remove("on");
  $("emptyHint").hidden = false; drop.classList.remove("compact"); setDropText(); updateDock();
}

$("clearBtn").addEventListener("click", async () => {
  if (S.busy) return;
  for (const r of S.results) files.del("pdf:" + r.id).catch(() => {});
  await kv.del("pending");
  hideNotice(); resetList();
});

// ───────────────────────────── скачивание ─────────────────────────────
async function download() {
  try {
    if (S.mode === "own") {
      const blob = await files.get("own-workbook");
      downloadBlob(blob, S.own.name);
      S.own.exportedAppended = S.own.appended;
      await saveOwn();
    } else {
      downloadBlob(await exportOurTable(S.table), "Счета.xlsx");
      S.table.exportedRows = S.table.rows.length;
      await saveTable();
    }
    renderChrome();
  } catch (e) {
    showNotice("bad", "Не получилось скачать: " + e.message);
  }
}
$("dlBtn").addEventListener("click", download);
$("remindDl").addEventListener("click", download);

// ───────────────────────────── настройки ─────────────────────────────
function renderSettings() {
  const k = S.settings.key || "";
  $("keyMasked").textContent = k ? `Сейчас: ${k.slice(0, 8)}…${k.slice(-4)}` : "Ключ не задан";
  $("modelSelect").replaceChildren(...Object.entries(MODELS).map(([id, m]) => h("option", { value: id, selected: id === (S.settings.model || DEFAULT_MODEL) }, m.label)));
  $("tableInfo").textContent = S.mode === "own"
    ? `Ваша таблица «${S.own.name}», лист «${S.own.sheetName}». Добавлено с сайта: ${S.own.appended}.`
    : `Таблица сайта «Счета.xlsx», строк: ${S.table.rows.length}.`;
  $("dlCopy").hidden = S.mode !== "own";
  const j = S.table.journal || [];
  const ym = (t) => (t || "").slice(3, 10);
  const month = ym(stamp());
  const spentMonth = j.filter((x) => ym(x.time) === month).reduce((s, x) => s + (+x.cost || 0), 0);
  const spentAll = j.reduce((s, x) => s + (+x.cost || 0), 0);
  $("spent").textContent = `Потрачено на распознавание: в этом месяце ≈ $${spentMonth.toFixed(3)}, всего ≈ $${spentAll.toFixed(3)}.`;
  $("journalTable").replaceChildren(...(j.length ? j.slice(-60).reverse().map((x) => h("tr", {}, h("td", {}, x.time), h("td", {}, x.file), h("td", {}, x.result), h("td", {}, x.cost ? `$${(+x.cost).toFixed(3)}` : "")))
    : [h("tr", {}, h("td", { class: "muted" }, "Пока пусто"))]));
  $("wipeBtn").textContent = "Стереть всё из этого браузера";
  $("wipeBtn").dataset.armed = "";
}
$("settingsBtn").addEventListener("click", () => { renderSettings(); $("settings").showModal(); });
$("settingsClose").addEventListener("click", () => $("settings").close());
$("keySave2").addEventListener("click", () => saveKey($("keyInput2"), $("keyErr2"), renderSettings));
$("modelSelect").addEventListener("change", async (e) => { S.settings.model = e.target.value; await kv.set("settings", S.settings); });
$("dlBtn2").addEventListener("click", download);
$("dlCopy").addEventListener("click", async () => downloadBlob(await exportOurTable(S.table), "Счета (журнал сайта).xlsx"));
$("changeTable").addEventListener("click", () => { $("settings").close(); showChoose(true); });
$("wipeBtn").addEventListener("click", async (e) => {
  if (!e.target.dataset.armed) { e.target.dataset.armed = "1"; e.target.textContent = "Нажмите ещё раз — ключ, таблица и журнал будут стёрты"; return; }
  await wipeAll();
  location.reload();
});

// ───────────────────────────── запуск ─────────────────────────────
async function route() {
  if (!S.settings.key) { show("viewKey"); renderChrome(); return; }
  if (!S.mode || (S.mode === "own" && !S.own)) { showChoose(false); renderChrome(); return; }
  show("viewWork");
  renderChrome();
}

async function init() {
  try {
    S.settings = (await kv.get("settings")) || {};
    S.mode = (await kv.get("mode")) || null;
    S.table = (await kv.get("table")) || emptyTable();
    S.own = (await kv.get("own")) || null;
    const pending = await kv.get("pending");
    await route();
    if (pending?.results?.length && !$("viewWork").hidden) {
      S.include = pending.include || {};
      S.edits = pending.edits || {};
      pending.results.forEach((r) => addResult(r, true));
      $("progress").classList.add("on");
      $("progressText").textContent = "Восстановлены счета, которые ещё не добавлены в таблицу.";
    }
  } catch (e) {
    document.querySelector("main").prepend(h("div", { class: "notice on bad" },
      "Браузер не даёт сайту сохранять данные (возможно, приватный режим). Откройте сайт в обычном окне. " + e.message));
  }
}
init();
