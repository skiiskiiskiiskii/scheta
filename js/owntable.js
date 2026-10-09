// Режим «Продолжить мою таблицу»: чтение чужого .xlsx, поиск заголовков, сопоставление колонок
// и аккуратное дописывание строк. Файл НЕ пересобирается: меняется только XML нужного листа
// (плюс диапазоны фильтра/«умной таблицы»), всё остальное — формулы, другие листы, графики — остаётся как было.
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

// Поля, которые сайт умеет заполнять (ключ → подпись в списке)
export const FIELDS = [
  ["Покупатель", "Покупатель"], ["Номер заказа", "Номер заказа"], ["Перевозчик", "Перевозчик (машина)"],
  ["Название", "Название товара (или «N поз.»)"], ["Позиций", "Число позиций в счёте"], ["Сумма", "Сумма (итог с НДС)"], ["Сумма без НДС", "Сумма без НДС"], ["НДС", "НДС (сумма)"],
  ["Литры", "Литры"], ["Количество", "Количество (как в счёте)"], ["Ед.", "Единица (л, ks, шт…)"],
  ["Цена за л", "Цена за литр"], ["Цена за ед.", "Цена за единицу"], ["Дата", "Дата счёта"], ["Дата отгрузки", "Дата отгрузки"],
  ["Валюта", "Валюта"], ["Поставщик", "Поставщик"], ["№ счёта", "№ счёта"],
  ["_status", "Статус проверки (OK / Проверить…)"], ["_notes", "Что проверить (пояснения)"], ["_file", "Имя PDF-файла"],
];
export const FIELD_LABEL = Object.fromEntries(FIELDS);

const parseXml = (s) => new DOMParser().parseFromString(s, "application/xml");
const serialize = (doc) => {
  const s = new XMLSerializer().serializeToString(doc);
  return s.startsWith("<?xml") ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
};
const els = (node, tag) => [...node.getElementsByTagNameNS(NS, tag)];
const textOf = (node) => els(node, "t").map((t) => t.textContent).join("");

export const colToIdx = (c) => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const idxToCol = (i) => { let s = ""; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const splitRef = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? [m[1], +m[2]] : [null, null]; };

function resolveTarget(base, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/").slice(0, -1);
  for (const p of target.split("/")) { if (p === "..") parts.pop(); else if (p !== ".") parts.push(p); }
  return parts.join("/");
}
const relsPath = (path) => { const p = path.split("/"); const f = p.pop(); return [...p, "_rels", f + ".rels"].join("/"); };

// ── форматы дат ──
const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);
function isDateFormat(id, code) {
  if (BUILTIN_DATE_IDS.has(id)) return true;
  if (!code) return false;
  const c = code.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").replace(/\\./g, "");
  return /[dmy]/i.test(c) && !/^[#0.,\s%]*$/.test(c);
}

/** Открывает .xlsx и возвращает книгу с удобными методами. */
export async function openWorkbook(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const wbXml = zip.file("xl/workbook.xml");
  if (!wbXml) throw new Error("это не файл Excel (.xlsx). Старый формат .xls нужно пересохранить в Excel как .xlsx");
  const wb = parseXml(await wbXml.async("string"));
  const rels = parseXml(await zip.file("xl/_rels/workbook.xml.rels").async("string"));
  const relMap = Object.fromEntries([...rels.getElementsByTagNameNS(NS_PKG_REL, "Relationship")].map((r) => [r.getAttribute("Id"), r.getAttribute("Target")]));
  const sheets = els(wb, "sheet").map((s) => ({
    name: s.getAttribute("name"), hidden: ["hidden", "veryHidden"].includes(s.getAttribute("state")),
    path: resolveTarget("xl/workbook.xml", relMap[s.getAttributeNS(NS_REL, "id")] || ""),
  })).filter((s) => s.path && zip.file(s.path));

  const shared = [];
  if (zip.file("xl/sharedStrings.xml")) {
    const ss = parseXml(await zip.file("xl/sharedStrings.xml").async("string"));
    for (const si of els(ss, "si")) shared.push(textOf(si));
  }
  const dateStyles = new Set();
  if (zip.file("xl/styles.xml")) {
    const st = parseXml(await zip.file("xl/styles.xml").async("string"));
    const custom = Object.fromEntries(els(st, "numFmt").map((n) => [+n.getAttribute("numFmtId"), n.getAttribute("formatCode")]));
    const xfs = els(st, "cellXfs")[0];
    if (xfs) els(xfs, "xf").filter((x) => x.parentNode === xfs).forEach((x, i) => {
      const id = +(x.getAttribute("numFmtId") || 0);
      if (isDateFormat(id, custom[id])) dateStyles.add(String(i));
    });
  }
  return { zip, sheets, shared, dateStyles };
}

/** Читает лист: Map(номер строки → Map(колонка → {v, t, s, f})) */
export async function readSheet(book, path) {
  const doc = parseXml(await book.zip.file(path).async("string"));
  const rows = new Map();
  let maxRow = 0;
  for (const r of els(doc, "row")) {
    const rn = +r.getAttribute("r");
    const cells = new Map();
    for (const c of els(r, "c")) {
      const [col] = splitRef(c.getAttribute("r") || "");
      if (!col) continue;
      const t = c.getAttribute("t"), s = c.getAttribute("s"), vEl = els(c, "v")[0], f = els(c, "f")[0];
      let v = null;
      if (t === "s") v = book.shared[+(vEl?.textContent)] ?? null;
      else if (t === "inlineStr") v = textOf(c);
      else if (vEl) v = t === "str" || t === "e" ? vEl.textContent : t === "b" ? vEl.textContent === "1" : Number(vEl.textContent);
      if (v === "" ) v = null;
      cells.set(col, { v, t, s, f: f ? f.textContent || "(общая формула)" : null, isDate: typeof v === "number" && book.dateStyles.has(s) });
    }
    rows.set(rn, cells);
    if (cells.size) maxRow = Math.max(maxRow, rn);
  }
  return { rows, maxRow };
}

// ── сопоставление колонок по названию ──
export const normalize = (s) => String(s ?? "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "")
  .replace(/ё/g, "е").replace(/[^\p{L}\p{N}№%/()., ]+/gu, " ").replace(/\s+/g, " ").trim();

const PRICE = /цена|cena|price|preis|cijena/;
const TOTALISH = /итог|всего|сумм|celkem|castka|total|amount|ukupno|iznos/;
const RULES = [
  ["Покупатель", /покупател|заказчик|клиент|buyer|customer|odberatel|kupac|zakaznik|kunde|abnehmer/],
  ["Номер заказа", /заказ|order|objednav|narudzb|bestell|^po\b|purchase/],
  ["Перевозчик", /перевоз|carrier|dopravce|prevoznik|spedit|truck|vehicle|машин|авто/],
  ["Дата отгрузки", /отгруз|поставк|duzp|zdanitel|uskut|dodan|dodaci|deliver|shipment|shipping|lieferdat|promet/],
  ["№ счёта", /№|номер|^(no|nr|n)\.?( |$)|cislo|invoice ?(no|number|#)|^faktura$|doklad|rechnungs?n|^inv(oice)?$|^broj/],
  ["Дата", /дата|datum|^date|date$|vystaven|ausgestellt/],
  ["Поставщик", /поставщ|продав|контрагент|dodavatel|supplier|vendor|seller|lieferant|prodavac|dobavljac/],
  ["Валюта", /валют|^mena$|currency|wahrung|valuta/],
  ["Ед.", /^ед\b|ед\.? ?изм|^mj$|^m\.j\.?$|jednotk|^unit$|einheit|^jm$/],
  ["Сумма без НДС", /без ндс|bez dph|zaklad|netto|\bnet\b|без налог|excl|bez pdv/],
  ["НДС", /^ндс\b|сумма ндс|^dph\b|^vat\b|mwst|tax amount|^pdv\b/],
  ["Сумма", /сумм|итог|всего|стоимост|celkem|castka|amount|total|brutto|s dph|с ндс|gross|^sum|iznos|ukupno/],
  ["Литры", /литр|^л\.?$|[ ,(]л\)?$|объ[её]м|litr|liter|litre|objem|volume|^l$|\(l\)|zapremina/],
  ["Количество", /кол|количеств|mnozstvi|qty|quantity|pocet|^ks$|^шт|menge|pcs|kolicina/],
  ["Название", /назван|наимен|товар|продук|позици|номенклат|polozk|nazev|zbozi|popis|oznacen|product|item|description|artikel|bezeichn|^name$|материал|naziv|artikl/],
  ["_notes", /примеч|коммент|poznam|note|comment|bemerk|napomena/],
  ["_status", /статус|status|проверк|kontrol/],
];

export function fieldForHeader(text) {
  const h = normalize(text);
  if (!h) return null;
  if (PRICE.test(h) && !TOTALISH.test(h)) return /за л|\/ ?l\b|\/ ?л\b|za l|per l|литр|liter|litr/.test(h) ? "Цена за л" : "Цена за ед.";
  for (const [field, re] of RULES) if (re.test(h)) return field;
  return null;
}

/** Строка заголовков: та, где больше всего узнаваемых названий колонок. */
export function detectHeader(sheet) {
  let best = null, bestScore = 0;
  for (let r = 1; r <= Math.min(sheet.maxRow, 30); r++) {
    const cells = sheet.rows.get(r);
    if (!cells) continue;
    const texts = [...cells.values()].filter((c) => typeof c.v === "string");
    const score = texts.filter((c) => fieldForHeader(c.v)).length;
    if (score > bestScore) { best = r; bestScore = score; }
  }
  if (best) return best;
  for (let r = 1; r <= Math.min(sheet.maxRow, 30); r++) {
    const cells = sheet.rows.get(r);
    if (cells && [...cells.values()].filter((c) => typeof c.v === "string").length >= 2) return r;
  }
  return 1;
}

const serialToIso = (n) => new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10);
const isoToSerial = (iso) => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };

export function displayValue(c) {
  if (!c || c.v === null) return "";
  if (c.isDate) return serialToIso(c.v).split("-").reverse().join(".");
  if (typeof c.v === "number") return c.v.toLocaleString("ru-RU", { maximumFractionDigits: 4 });
  return String(c.v);
}

/** Последняя заполненная строка под заголовком (по колонкам, у которых есть заголовок). */
export function lastDataRow(sheet, headerRow) {
  const cols = [...(sheet.rows.get(headerRow)?.keys() || [])];
  let last = headerRow;
  for (const [rn, cells] of sheet.rows) {
    if (rn <= headerRow) continue;
    if (cols.some((c) => cells.get(c)?.v !== null && cells.get(c)?.v !== undefined)) last = Math.max(last, rn);
  }
  return last;
}

/** Колонки для экрана сопоставления: буква, заголовок, примеры, предложенное поле. */
export function describeColumns(sheet, headerRow) {
  const head = sheet.rows.get(headerRow) || new Map();
  const last = lastDataRow(sheet, headerRow);
  const out = [];
  for (const [col, c] of [...head.entries()].sort((a, b) => colToIdx(a[0]) - colToIdx(b[0]))) {
    if (c.v === null) continue;
    const samples = [];
    for (let r = last; r > headerRow && samples.length < 3; r--) {
      const v = displayValue(sheet.rows.get(r)?.get(col));
      if (v) samples.push(v);
    }
    out.push({ col, header: String(c.v), samples, field: fieldForHeader(c.v) });
  }
  // одно поле — одна колонка: оставляем первую
  const used = new Set();
  for (const c of out) { if (c.field && used.has(c.field)) c.field = null; else if (c.field) used.add(c.field); }
  return out;
}

/** Похоже ли, что последняя строка — «Итого»? */
export function totalsRowHint(sheet, headerRow) {
  const last = lastDataRow(sheet, headerRow);
  const cells = sheet.rows.get(last);
  if (!cells || last === headerRow) return null;
  const hasTotalWord = [...cells.values()].some((c) => typeof c.v === "string" && /итог|всего|celkem|total|ukupno|summe/i.test(c.v));
  const hasSumFormula = [...cells.values()].some((c) => c.f && /SUM|СУММ/i.test(c.f));
  return hasTotalWord || hasSumFormula ? last : null;
}

/** Что уже есть в таблице брата — для предупреждений о дублях. */
export function existingIndex(sheet, headerRow, mapping) {
  const colOf = Object.fromEntries(Object.entries(mapping).filter(([, f]) => f).map(([c, f]) => [f, c]));
  const numbers = {}, fuzzy = [], orders = [];
  for (const [rn, cells] of sheet.rows) {
    if (rn <= headerRow) continue;
    const get = (f) => (colOf[f] ? cells.get(colOf[f]) : null);
    const n = get("№ счёта")?.v;
    if (n !== null && n !== undefined) numbers[String(n).toLowerCase().replace(/[\W_]/gu, "")] = rn;
    const o = get("Номер заказа")?.v;
    if (o !== null && o !== undefined && String(o).trim()) orders.push({ row: rn, text: String(o).toLowerCase().replace(/[\W_]/gu, "") });
    const d = get("Дата")?.isDate ? serialToIso(get("Дата").v) : null;
    const sum = typeof get("Сумма")?.v === "number" ? get("Сумма").v : typeof get("Сумма без НДС")?.v === "number" ? get("Сумма без НДС").v : null;
    const lit = typeof get("Литры")?.v === "number" ? get("Литры").v : null;
    if (sum !== null || lit !== null) fuzzy.push({ row: rn, date: d, sum, lit });
  }
  return { numbers, fuzzy, orders };
}

/** Предупреждения о возможных дублях со старыми строками таблицы брата. */
export function similarRows(index, invNumber, row) {
  const out = [];
  const n = invNumber ? String(invNumber).toLowerCase().replace(/[\W_]/gu, "") : null;
  if (n && index.numbers[n]) out.push(`счёт № ${invNumber} уже встречается в вашей таблице (строка ${index.numbers[n]}) — проверьте, не дубль ли`);
  const v = row.values;
  const ord = v["Номер заказа"] ? String(v["Номер заказа"]).toLowerCase().replace(/[\W_]/gu, "") : null;
  const oh = ord && ord.length >= 4 ? (index.orders || []).find((o) => o.text.includes(ord)) : null;
  if (oh) out.push(`заказ ${v["Номер заказа"]} уже встречается в вашей таблице (строка ${oh.row}) — проверьте, не дубль ли`);
  const near = (a, b) => a !== null && b !== null && b !== undefined && Math.abs(a - b) < 0.01;
  const hit = index.fuzzy.find((f) => (f.date ? f.date === v["Дата"] : true) &&
    ((near(f.sum, v["Сумма"]) || near(f.sum, v["Сумма без НДС"])) && (f.lit === null || near(f.lit, v["Литры"]))) && (f.date || f.lit !== null));
  if (hit) out.push(`похожая строка уже есть в вашей таблице (строка ${hit.row}): та же ${hit.date ? "дата и " : ""}сумма — проверьте, не дубль ли`);
  return out;
}

// ── дописывание строк ──
export function rowValue(field, row, file, notes) {
  if (field === "_status") return row.status;
  if (field === "_notes") return notes;
  if (field === "_file") return file;
  return row.values[field] ?? null;
}

function makeCell(doc, ref, style, value, asDate, dateAsText) {
  const c = doc.createElementNS(NS, "c");
  c.setAttribute("r", ref);
  if (style) c.setAttribute("s", style);
  if (value === null || value === undefined || value === "") return c;
  if (asDate && typeof value === "string" && /^\d{4}-\d\d-\d\d$/.test(value)) {
    if (dateAsText) value = value.split("-").reverse().join(".");
    else { const v = doc.createElementNS(NS, "v"); v.textContent = String(isoToSerial(value)); c.append(v); return c; }
  }
  if (typeof value === "number") { const v = doc.createElementNS(NS, "v"); v.textContent = String(value); c.append(v); return c; }
  c.setAttribute("t", "inlineStr");
  const is = doc.createElementNS(NS, "is"), t = doc.createElementNS(NS, "t");
  t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.textContent = String(value);
  is.append(t); c.append(is);
  return c;
}

function extendRef(ref, oldLast, newLast) {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref || "");
  if (!m) return ref;
  return +m[4] >= oldLast - 1 && +m[4] < newLast ? `${m[1]}${m[2]}:${m[3]}${newLast}` : ref;
}

/**
 * Дописывает строки в лист. rows: [{row, file, notes}], mapping: {колонка: поле}.
 * Возвращает {first, last}. Стиль каждой новой ячейки берётся у последней заполненной строки той же колонки.
 */
export async function appendRows(book, sheetPath, headerRow, mapping, items) {
  const { zip } = book;
  const sheet = await readSheet(book, sheetPath);
  const last = lastDataRow(sheet, headerRow);
  const doc = parseXml(await zip.file(sheetPath).async("string"));
  const sheetData = els(doc, "sheetData")[0];
  const rowEls = new Map(els(sheetData, "row").map((r) => [+r.getAttribute("r"), r]));
  const cols = Object.entries(mapping).filter(([, f]) => f).sort((a, b) => colToIdx(a[0]) - colToIdx(b[0]));

  // образец оформления: последняя непустая ячейка в каждой колонке (до 50 строк вверх)
  const sample = {};
  for (const [col] of cols) {
    for (let r = last; r > headerRow && r > last - 50; r--) {
      const c = sheet.rows.get(r)?.get(col);
      if (c && c.v !== null) { sample[col] = c; break; }
    }
  }
  const lastRowEl = rowEls.get(last);
  items.forEach((item, i) => {
    const rn = last + 1 + i;
    let rowEl = rowEls.get(rn);
    if (!rowEl) {
      rowEl = doc.createElementNS(NS, "row");
      rowEl.setAttribute("r", String(rn));
      if (lastRowEl?.getAttribute("ht") && lastRowEl.getAttribute("customHeight")) {
        rowEl.setAttribute("ht", lastRowEl.getAttribute("ht")); rowEl.setAttribute("customHeight", "1");
      }
      const after = [...rowEls.keys()].filter((k) => k > rn).sort((a, b) => a - b)[0];
      sheetData.insertBefore(rowEl, after ? rowEls.get(after) : null);
      rowEls.set(rn, rowEl);
    }
    rowEl.removeAttribute("spans");
    for (const [col, field] of cols) {
      const ref = col + rn;
      const s = sample[col];
      const isDateField = field === "Дата" || field === "Дата отгрузки";
      const cell = makeCell(doc, ref, s?.s || els(rowEl, "c").find((c) => c.getAttribute("r") === ref)?.getAttribute("s"),
        rowValue(field, item.row, item.file, item.notes), isDateField, isDateField && (!s || !s.isDate));
      const existing = els(rowEl, "c").find((c) => c.getAttribute("r") === ref);
      if (existing) rowEl.replaceChild(cell, existing);
      else {
        const next = els(rowEl, "c").find((c) => colToIdx(splitRef(c.getAttribute("r"))[0]) > colToIdx(col));
        rowEl.insertBefore(cell, next || null);
      }
    }
  });
  const newLast = last + items.length;

  // диапазоны: размер листа, фильтр, «умная таблица»
  const dim = els(doc, "dimension")[0];
  if (dim) {
    const m = /^([A-Z]+\d+)(?::([A-Z]+)(\d+))?$/.exec(dim.getAttribute("ref") || "");
    if (m) {
      const endCol = m[2] ? (colToIdx(m[2]) >= colToIdx(cols.at(-1)?.[0] || "A") ? m[2] : cols.at(-1)[0]) : cols.at(-1)?.[0] || "A";
      dim.setAttribute("ref", `${m[1]}:${endCol}${Math.max(+(m[3] || 1), newLast)}`);
    }
  }
  const af = els(doc, "autoFilter")[0];
  if (af) af.setAttribute("ref", extendRef(af.getAttribute("ref"), last, newLast));
  zip.file(sheetPath, serialize(doc));

  const rp = relsPath(sheetPath);
  if (zip.file(rp)) {
    const rels = parseXml(await zip.file(rp).async("string"));
    for (const r of rels.getElementsByTagNameNS(NS_PKG_REL, "Relationship")) {
      if (!/\/table$/.test(r.getAttribute("Type") || "")) continue;
      const tp = resolveTarget(sheetPath, r.getAttribute("Target"));
      if (!zip.file(tp)) continue;
      const t = parseXml(await zip.file(tp).async("string"));
      const tab = t.documentElement;
      const newRef = extendRef(tab.getAttribute("ref"), last, newLast);
      if (newRef !== tab.getAttribute("ref")) {
        tab.setAttribute("ref", newRef);
        const taf = els(t, "autoFilter")[0];
        if (taf) taf.setAttribute("ref", extendRef(taf.getAttribute("ref"), last, newLast));
        zip.file(tp, serialize(t));
      }
    }
  }
  return { first: last + 1, last: newLast };
}

/** Лист из «таблицы значений» (как отдаёт Google Таблица): даты дд.мм.гггг и числа распознаются. */
export function sheetFromValues(values) {
  const rows = new Map();
  let maxRow = 0;
  values.forEach((row, i) => {
    const cells = new Map();
    row.forEach((raw, j) => {
      const t = String(raw ?? "").trim();
      if (!t) return;
      let v = t, isDate = false;
      const d = t.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
      const numText = t.replace(/[€$£₽]|Kč|CZK|EUR|RSD|руб\.?|\u00a0|\u202f|\s/gi, "");
      if (d) { v = (Date.UTC(+d[3], +d[2] - 1, +d[1]) - Date.UTC(1899, 11, 30)) / 86400000; isDate = true; }
      else if (/^-?\d[\d.,]*$/.test(numText)) {
        let n = numText;
        if (n.includes(",") && n.includes(".")) { const dec = n.lastIndexOf(",") > n.lastIndexOf(".") ? "," : "."; n = n.split(dec === "," ? "." : ",").join("").replace(dec, "."); }
        else if (n.includes(",")) n = n.split(",").length > 2 ? n.split(",").join("") : n.replace(",", ".");
        if (Number.isFinite(Number(n))) v = Number(n);
      }
      cells.set(idxToCol(j + 1), { v, t: null, s: null, f: null, isDate });
    });
    rows.set(i + 1, cells);
    if (cells.size) maxRow = i + 1;
  });
  return { rows, maxRow };
}

export async function toBlob(book) {
  return book.zip.generateAsync({ type: "blob", mimeType: XLSX_MIME, compression: "DEFLATE" });
}
