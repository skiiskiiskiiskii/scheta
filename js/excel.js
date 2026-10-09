// Наша таблица «Счета.xlsx»: создание (скачать) и чтение обратно (восстановить).
// Формат совпадает с Python-версией: лист «Счета», «Журнал» и скрытый «_служебное».
import { OK, WARN, ERROR, EDITED } from "./core.js";

export const COLUMNS = [
  ["Название", 34], ["Сумма", 14], ["Литры", 11], ["Дата", 11], ["Дата отгрузки", 14], ["Валюта", 10], ["Цена за л", 12],
  ["Поставщик", 28], ["№ счёта", 13], ["Статус", 12], ["Что проверить", 55], ["Файл", 34],
];
const FILL = { [OK]: "FFE2F0D9", [WARN]: "FFFFF2CC", [ERROR]: "FFF8CBAD", [EDITED]: "FFDDEBF7" };
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const isoToDate = (iso) => { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const dateToIso = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;

export function rowNotes(row) {
  const msgs = row.messages.map(([, m]) => m);
  if (row.edits?.length) msgs.push("исправлено вручную: " + row.edits.join(", "));
  return msgs.join("; ") || null;
}

export async function exportOurTable(table) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Счета → Excel";
  const ws = wb.addWorksheet("Счета", { views: [{ state: "frozen", ySplit: 1 }] });
  ws.columns = COLUMNS.map(([header, width]) => ({ header, width }));
  const h = ws.getRow(1);
  h.height = 30;
  h.eachCell((c) => {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF44546A" } };
    c.alignment = { vertical: "middle", wrapText: true };
  });
  table.rows.forEach((row, i) => {
    const v = row.values;
    const r = ws.getRow(i + 2);
    const vals = [v["Название"], v["Сумма"], v["Литры"], v["Дата"] ? isoToDate(v["Дата"]) : null,
      v["Дата отгрузки"] ? isoToDate(v["Дата отгрузки"]) : null, v["Валюта"], v["Цена за л"], v["Поставщик"], v["№ счёта"],
      row.status, rowNotes(row), row.link ? { text: row.file, hyperlink: row.link } : row.file];
    vals.forEach((x, j) => {
      const c = r.getCell(j + 1);
      c.value = x ?? null;
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: FILL[row.status] || FILL[OK] } };
      c.alignment = { vertical: "top", wrapText: [0, 7, 10].includes(j) };
    });
    r.getCell(2).numFmt = "#,##0.00";
    r.getCell(3).numFmt = Number.isInteger(v["Литры"]) ? "#,##0" : "#,##0.00#";
    r.getCell(4).numFmt = r.getCell(5).numFmt = "dd.mm.yyyy";
    r.getCell(7).numFmt = "#,##0.00##";
    r.getCell(10).font = { bold: true };
    if (row.link) r.getCell(12).font = { color: { argb: "FF0563C1" }, underline: true };
  });
  if (table.rows.length) ws.autoFilter = { from: "A1", to: `L${table.rows.length + 1}` };

  const log = wb.addWorksheet("Журнал");
  log.columns = [["Время", 17], ["Файл", 40], ["Результат", 70], ["Строк добавлено", 10], ["Модель", 18], ["Токены (вход/выход)", 18], ["Стоимость, $", 12]]
    .map(([header, width]) => ({ header, width }));
  log.getRow(1).font = { bold: true };
  for (const j of table.journal) log.addRow([j.time, j.file, j.result, j.rows ?? 0, j.model || "", j.tokens || "", j.cost || 0]);

  const svc = wb.addWorksheet("_служебное", { state: "hidden" });
  svc.addRow(["ключ счёта", "sha256 файла", "строка в «Счета»", "файл"]);
  table.rows.forEach((row, i) => svc.addRow([row.key, row.sha, i + 2, row.file]));
  // дубли, которые помним без строк (например, из прошлых загрузок)
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
  if (!ws || !svc || ws.getRow(1).getCell(1).value !== "Название") return null;
  const header = {};
  ws.getRow(1).eachCell((c, col) => { header[c.value] = col; });
  const get = (r, name) => (header[name] ? cellText(r.getCell(header[name]).value) : null);
  const rows = [];
  for (let i = 2; i <= ws.rowCount; i++) {
    const r = ws.getRow(i);
    if (!get(r, "Название") && !get(r, "Сумма")) continue;
    const status = [OK, WARN, ERROR, EDITED].includes(get(r, "Статус")) ? get(r, "Статус") : OK;
    const notes = get(r, "Что проверить");
    const numOrNull = (x) => (typeof x === "number" ? x : x === null ? null : Number(String(x).replace(/\s/g, "").replace(",", ".")) || null);
    rows.push({
      values: { "Название": get(r, "Название"), "Сумма": numOrNull(get(r, "Сумма")), "Литры": numOrNull(get(r, "Литры")),
        "Дата": get(r, "Дата"), "Дата отгрузки": get(r, "Дата отгрузки"), "Валюта": get(r, "Валюта"),
        "Цена за л": numOrNull(get(r, "Цена за л")), "Поставщик": get(r, "Поставщик"), "№ счёта": get(r, "№ счёта") != null ? String(get(r, "№ счёта")) : null },
      status, messages: notes ? [[status === OK ? "info" : status, String(notes)]] : [], edits: [],
      file: get(r, "Файл"), key: null, sha: null, restored: true,
    });
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
