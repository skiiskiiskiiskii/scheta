// Наша таблица «Счета.xlsx»: создание (скачать) и чтение обратно (восстановить).
// Формат совпадает с Python-версией: лист «Счета», «Журнал» и скрытый «_служебное».
import { OK, WARN, ERROR, EDITED } from "./core.js";

export const COLUMNS_ITEMS = [
  ["Название", 34], ["Сумма", 14], ["Литры", 11], ["Дата", 11], ["Дата отгрузки", 14], ["Валюта", 10], ["Цена за л", 12],
  ["Поставщик", 28], ["№ счёта", 13], ["Статус", 12], ["Что проверить", 55], ["Файл", 34],
];
export const COLUMNS_SUMMARY = [
  ["Дата отгрузки", 13], ["№ счёта", 14], ["Номер заказа", 14], ["Перевозчик", 16], ["Покупатель", 30], ["Сумма", 14],
  ["Валюта", 8], ["Литры", 11], ["Позиций", 9], ["Поставщик", 24], ["Дата", 11], ["Статус", 12], ["Что проверить", 55], ["Файл", 34],
];
export const COLUMNS = COLUMNS_ITEMS;
const FILL = { [OK]: "FFE2F0D9", [WARN]: "FFFFF2CC", [ERROR]: "FFF8CBAD", [EDITED]: "FFDDEBF7" };
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DATE_COLS = new Set(["Дата", "Дата отгрузки"]);
const NUM_FMT = { "Сумма": "#,##0.00", "Сумма без НДС": "#,##0.00", "НДС": "#,##0.00", "Цена за л": "#,##0.00##", "Позиций": "0" };

const isoToDate = (iso) => { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const dateToIso = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;

export function rowNotes(row) {
  const msgs = row.messages.map(([, m]) => m);
  if (row.edits?.length) msgs.push("исправлено вручную: " + row.edits.join(", "));
  return msgs.join("; ") || null;
}

export function cellValue(name, row) {
  if (name === "Статус") return row.status;
  if (name === "Что проверить") return rowNotes(row);
  if (name === "Файл") return row.file;
  const v = row.values[name];
  return DATE_COLS.has(name) && v ? isoToDate(v) : v ?? null;
}

export async function exportOurTable(table, layout = "summary") {
  const columns = layout === "items" ? COLUMNS_ITEMS : COLUMNS_SUMMARY;
  const wb = new ExcelJS.Workbook();
  wb.creator = "Счета → Excel";
  const ws = wb.addWorksheet("Счета", { views: [{ state: "frozen", ySplit: 1 }] });
  ws.columns = columns.map(([header, width]) => ({ header, width }));
  const h = ws.getRow(1);
  h.height = 30;
  h.eachCell((c) => {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF44546A" } };
    c.alignment = { vertical: "middle", wrapText: true };
  });
  table.rows.forEach((row, i) => {
    const r = ws.getRow(i + 2);
    columns.forEach(([name], j) => {
      const c = r.getCell(j + 1);
      const v = cellValue(name, row);
      c.value = name === "Файл" && row.link ? { text: row.file, hyperlink: row.link } : v;
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: FILL[row.status] || FILL[OK] } };
      c.alignment = { vertical: "top", wrapText: ["Название", "Поставщик", "Покупатель", "Что проверить"].includes(name) };
      if (DATE_COLS.has(name)) c.numFmt = "dd.mm.yyyy";
      else if (name === "Литры") c.numFmt = Number.isInteger(row.values["Литры"]) ? "#,##0" : "#,##0.00#";
      else if (NUM_FMT[name]) c.numFmt = NUM_FMT[name];
      if (name === "Статус") c.font = { bold: true };
    });
  });
  if (table.rows.length) ws.autoFilter = { from: "A1", to: { row: table.rows.length + 1, column: columns.length } };

  const log = wb.addWorksheet("Журнал");
  log.columns = [["Время", 17], ["Файл", 40], ["Результат", 70], ["Строк добавлено", 10], ["Модель", 18], ["Токены (вход/выход)", 18], ["Стоимость, $", 12]]
    .map(([header, width]) => ({ header, width }));
  log.getRow(1).font = { bold: true };
  for (const j of table.journal) log.addRow([j.time, j.file, j.result, j.rows ?? 0, j.model || "", j.tokens || "", j.cost || 0]);

  const svc = wb.addWorksheet("_служебное", { state: "hidden" });
  svc.addRow(["ключ счёта", "sha256 файла", "строка в «Счета»", "файл"]);
  table.rows.forEach((row, i) => svc.addRow([row.key, row.sha, i + 2, row.file]));
  return new Blob([await wb.xlsx.writeBuffer()], { type: XLSX_MIME });
}

function cellText(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return dateToIso(v);
  if (typeof v === "object") {
    if (v.richText) return v.richText.map((t) => t.text).join("");
    if ("text" in v) return typeof v.text === "object" ? cellText(v.text) : v.text;
    if ("result" in v) return v.result;
  }
  return v;
}

/** Это наша таблица? Если да — возвращает данные для восстановления, иначе null. */
export async function readOurTable(arrayBuffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  const ws = wb.getWorksheet("Счета"), svc = wb.getWorksheet("_служебное");
  if (!ws || !svc) return null;
  const header = {};
  ws.getRow(1).eachCell((c, col) => { header[cellText(c.value)] = col; });
  if (!header["Статус"] || !header["Что проверить"]) return null;
  const get = (r, name) => (header[name] ? cellText(r.getCell(header[name]).value) : null);
  const numOrNull = (x) => (typeof x === "number" ? x : x === null || x === undefined || x === "" ? null : Number(String(x).replace(/\s/g, "").replace(",", ".")) || null);
  const NUMERIC = new Set(["Сумма", "Литры", "Цена за л", "Позиций", "Сумма без НДС", "НДС"]);
  const rows = [];
  for (let i = 2; i <= ws.rowCount; i++) {
    const r = ws.getRow(i);
    const values = {};
    for (const name of Object.keys(header)) {
      if (["Статус", "Что проверить", "Файл"].includes(name)) continue;
      const v = get(r, name);
      values[name] = NUMERIC.has(name) ? numOrNull(v) : v === null || v === undefined ? null : String(v);
    }
    if (Object.values(values).every((v) => v === null)) continue;
    const st = get(r, "Статус");
    const status = [OK, WARN, ERROR, EDITED].includes(st) ? st : OK;
    const notes = get(r, "Что проверить");
    rows.push({ values, status, messages: notes ? [[status === OK ? "info" : status, String(notes)]] : [], edits: [],
      file: get(r, "Файл"), key: null, sha: null, restored: true });
  }
  const keys = {}, hashes = {};
  for (let i = 2; i <= svc.rowCount; i++) {
    const r = svc.getRow(i);
    const [key, sha, rowNum, file] = [1, 2, 3, 4].map((c) => cellText(r.getCell(c).value));
    const idx = Number(rowNum) - 2;
    if (rows[idx]) { rows[idx].key = key || null; rows[idx].sha = sha || null; }
    const where = `в таблице (строка ${rowNum}, ${file})`;
    if (key) keys[key] = where;
    if (sha) hashes[sha] = where;
  }
  const journal = [];
  const log = wb.getWorksheet("Журнал");
  if (log) for (let i = 2; i <= log.rowCount; i++) {
    const r = log.getRow(i);
    const [time, file, result, n, model, tokens, cost] = [1, 2, 3, 4, 5, 6, 7].map((c) => cellText(r.getCell(c).value));
    if (time || file) journal.push({ time, file, result, rows: n, model, tokens, cost });
  }
  return { rows, keys, hashes, journal, exportedRows: rows.length };
}

export function downloadBlob(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
