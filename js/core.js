// Ядро: распознавание счёта через OpenAI и все проверки.
// Логика один в один с Python-версией (invoices.py), только работает прямо в браузере.
import * as pdfjs from "../vendor/pdf.min.mjs";

pdfjs.GlobalWorkerOptions.workerSrc = new URL("../vendor/pdf.worker.min.mjs", import.meta.url).href;

export const OK = "OK", WARN = "ПРОВЕРИТЬ", ERROR = "ОШИБКА", INFO = "info", EDITED = "ИСПРАВЛЕНО";
export const RANK = { [INFO]: 0, [OK]: 0, [EDITED]: 1, [WARN]: 2, [ERROR]: 3 };
export const MODELS = {
  "gpt-6.1-sol": { label: "gpt-6.1-sol — точнее (≈1–2 цента за счёт)", price: [2.0, 10.0] },
  "gpt-6-luna": { label: "gpt-6-luna — дешевле (доли цента)", price: [0.10, 0.50] },
};
export const DEFAULT_MODEL = "gpt-6.1-sol";

// ───────────────────────────── распознавание ─────────────────────────────
const NUM = { type: ["number", "null"] };
const STR = { type: ["string", "null"] };
export const SCHEMA = {
  type: "object", additionalProperties: false,
  required: ["is_invoice", "supplier_name", "supplier_tax_id", "invoice_number", "invoice_date", "shipment_date",
    "shipment_date_label", "currency", "items", "total_net", "total_vat", "total_gross", "uncertain_fields", "notes"],
  properties: {
    is_invoice: { type: "boolean" },
    supplier_name: STR, supplier_tax_id: STR, invoice_number: STR,
    invoice_date: { ...STR, description: "YYYY-MM-DD" },
    shipment_date: { ...STR, description: "YYYY-MM-DD" },
    shipment_date_label: STR,
    currency: { ...STR, description: "ISO 4217, e.g. RUB, CZK, EUR, RSD" },
    items: {
      type: "array", items: {
        type: "object", additionalProperties: false,
        required: ["name", "quantity", "unit", "package_volume_liters", "price_per_unit", "price_includes_vat",
          "amount_net", "vat_rate", "vat_amount", "amount_gross"],
        properties: {
          name: { type: "string" }, quantity: NUM, unit: STR,
          package_volume_liters: { ...NUM, description: "volume of ONE piece/package in litres, if quantity is in pieces and the volume is printed (e.g. '4L' in the name)" },
          price_per_unit: NUM, price_includes_vat: { type: ["boolean", "null"] },
          amount_net: NUM, vat_rate: NUM, vat_amount: NUM, amount_gross: NUM,
        },
      },
    },
    total_net: NUM, total_vat: NUM, total_gross: NUM,
    uncertain_fields: { type: "array", items: { type: "string" } },
    notes: STR,
  },
};

export const SYSTEM_PROMPT = `You extract data from supplier invoices for an accounting spreadsheet.
Documents may be in any language (Russian, Czech, Serbian, English, German, ...) and may be scans.

Accuracy rules — mistakes are costly, flagging is cheap:
- Transcribe every number EXACTLY as printed. Never calculate, round, fix or fill in a value.
  If a value is not printed, return null. If the printed arithmetic looks wrong, still return
  the printed values — the spreadsheet checks the arithmetic itself.
- Read numbers in the document's locale: "1 234,56" and "1.234,56" mean 1234.56; "1,234.56" means 1234.56.
- items: one entry per goods/service line. Never include subtotal/total/VAT recap rows as items.
  quantity and unit as printed ("л", "l", "L", "m3", "kg", "т", "шт", "ks", "pcs"...).
  package_volume_liters: if quantity is counted in pieces/packages (ks, шт, pcs, bal, kan...) and the
  volume of one package is printed (e.g. "Motorový olej 5W-30 4 l", "1L", "208 L sud", column "Objem"),
  give that volume in litres (ml → divide by 1000). Otherwise null. Never guess a volume that is not printed.
  price_includes_vat: true/false if the column header says so, else null.
  amount_net = line value without VAT, vat_amount = VAT of the line, amount_gross = line value with VAT.
  If the invoice has no VAT, put the line total into amount_gross and leave VAT fields null.
- supplier = seller/issuer (Продавец, Поставщик, Dodavatel, Prodavac, Seller), never the buyer.
  supplier_tax_id = seller's ИНН / IČO / DIČ / PIB / VAT ID as printed.
- Dates as YYYY-MM-DD. invoice_date = date of issue.
  shipment_date = date the goods were shipped/delivered/supplied, e.g. "Дата отгрузки",
  "Документ об отгрузке № … от …", "Datum uskutečnění zdanitelného plnění" (DUZP), "Datum dodání",
  "Datum prometa", "Delivery date", "Lieferdatum". shipment_date_label = the label you used. Null if absent.
- uncertain_fields: list every field you are not fully sure about (blurry, cut off, ambiguous),
  e.g. "items[0].quantity", "invoice_date". Be honest — a flagged field gets checked by a human.
- is_invoice = false if the document is not an invoice / bill with amounts.`;

const API = "https://api.openai.com/v1";

export function bytesToBase64(bytes) {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** Понятное описание ошибки API на русском. */
export function explainApiError(status, data, fallback) {
  const msg = data?.error?.message || fallback || "";
  const code = data?.error?.code || data?.error?.type || "";
  if (status === 401) return "ключ API неверный или удалён — проверьте его в настройках";
  if (status === 429 && /quota|billing|insufficient/i.test(msg + code))
    return "на счёте OpenAI закончились деньги — пополните баланс на platform.openai.com → Billing";
  if (status === 429) return "слишком много запросов подряд — подождите минуту и попробуйте ещё раз";
  if (status === 403) return `у ключа нет доступа к этой модели или функции (${msg.slice(0, 120)})`;
  if (status === 404) return `модель не найдена — выберите другую в настройках (${msg.slice(0, 120)})`;
  if (status >= 500) return "сервер OpenAI временно недоступен — попробуйте через пару минут";
  return msg ? `ответ OpenAI: ${msg.slice(0, 200)}` : `ошибка ${status}`;
}

async function callApi(path, key, body) {
  let r;
  try {
    r = await fetch(API + path, {
      method: body ? "POST" : "GET",
      headers: { Authorization: `Bearer ${key}`, ...(body ? { "Content-Type": "application/json" } : {}) },
      body: body ? JSON.stringify(body) : undefined,
    });
  } catch (e) {
    throw new Error("нет связи с OpenAI — проверьте интернет (или VPN/блокировщик) и попробуйте ещё раз");
  }
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(explainApiError(r.status, data, r.statusText));
  return data;
}

/** Что не так с ключом по виду (до запроса). null — выглядит нормально. */
export function keyProblem(key) {
  const k = (key || "").trim();
  if (!k) return "вставьте ключ";
  if (k.startsWith("sk-sk-")) return "в начале ключа лишнее «sk-» — должно быть ровно sk-proj-…";
  if (/\s/.test(k)) return "в ключе пробел или перенос строки — ключ должен быть одной строкой";
  if (!k.startsWith("sk-")) return "ключ OpenAI начинается с «sk-»";
  return null;
}

/** Бесплатная проверка ключа: список моделей. */
export async function validateKey(key) {
  const p = keyProblem(key);
  if (p) throw new Error(p);
  await callApi("/models", key.trim());
}

function outputText(data) {
  let text = "";
  for (const item of data.output || []) {
    if (item.type !== "message") continue;
    for (const c of item.content || []) {
      if (c.type === "output_text") text += c.text;
      if (c.type === "refusal") throw new Error("модель отказалась обрабатывать документ: " + c.refusal);
    }
  }
  return text;
}

function usageOf(model, u) {
  const [pin, pout] = MODELS[model]?.price || [0, 0];
  const tin = u?.input_tokens || 0, tout = u?.output_tokens || 0;
  return { model, input_tokens: tin, output_tokens: tout, cost_usd: Math.round((tin * pin + tout * pout) / 1e6 * 1e4) / 1e4 };
}

export async function extract(bytes, filename, key, model) {
  const data = await callApi("/responses", key, {
    model,
    instructions: SYSTEM_PROMPT,
    input: [{ role: "user", content: [
      { type: "input_file", filename, file_data: "data:application/pdf;base64," + bytesToBase64(bytes) },
      { type: "input_text", text: "Extract the invoice data." },
    ] }],
    text: { format: { type: "json_schema", name: "invoice", schema: SCHEMA, strict: true } },
    max_output_tokens: 16000,
  });
  if (data.status && data.status !== "completed")
    throw new Error(`ответ модели не завершён (${data.status}${data.incomplete_details ? ": " + JSON.stringify(data.incomplete_details) : ""})`);
  return { inv: JSON.parse(outputText(data)), usage: usageOf(model, data.usage) };
}

/** Небольшой запрос к модели с JSON-ответом (для сопоставления колонок). */
export async function askJson(key, model, instructions, text, schema, name = "answer") {
  const data = await callApi("/responses", key, {
    model, instructions,
    input: [{ role: "user", content: [{ type: "input_text", text }] }],
    text: { format: { type: "json_schema", name, schema, strict: true } },
    max_output_tokens: 4000,
  });
  return { json: JSON.parse(outputText(data)), usage: usageOf(model, data.usage) };
}

// ───────────────────────────── текст PDF ─────────────────────────────
export async function pdfText(bytes) {
  try {
    const task = pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false });
    const doc = await task.promise;
    const out = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const tc = await page.getTextContent();
      out.push(tc.items.map((i) => (i.str || "") + (i.hasEOL ? "\n" : " ")).join(""));
    }
    await task.destroy();
    return out.join("\n");
  } catch (e) {
    console.error("pdfText:", e);
    return "";
  }
}

export async function sha256(bytes) {
  const h = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

// ───────────────────────────── числа ─────────────────────────────
export const num = (x) => (x === null || x === undefined || x === "" || Number.isNaN(Number(x)) ? null : Number(x));
// округление до копеек «половина вверх» без ошибок двоичных дробей: 5.475 → 5.48
const r2 = (x) => Number(Math.round(Number(Number(x).toFixed(10) + "e2")) + "e-2");
const key4 = (x) => (Math.round(x * 10000) / 10000).toFixed(4);
const close = (a, b, tol) => Math.abs(a - b) <= tol + 1e-9;

export function fmt(x) {
  if (x === null || x === undefined) return "—";
  const neg = x < 0, a = Math.abs(x);
  const isInt = Math.abs(a - Math.round(a)) < 1e-9;
  const [i, d] = (isInt ? String(Math.round(a)) : a.toFixed(2)).split(".");
  return (neg ? "-" : "") + i.replace(/\B(?=(\d{3})+(?!\d))/g, " ") + (d ? "," + d : "");
}

function parseVariants(s) {
  s = s.replace(/'/g, "");
  const out = [];
  const add = (t) => { const n = Number(t); if (t !== "" && Number.isFinite(n)) out.push(n); };
  if (s.includes(",") && s.includes(".")) {
    const dec = s.lastIndexOf(",") > s.lastIndexOf(".") ? "," : ".";
    add(s.split(dec === "," ? "." : ",").join("").replace(dec, "."));
  } else if (s.includes(",") || s.includes(".")) {
    const sep = s.includes(",") ? "," : ".";
    if (s.split(sep).length > 2) add(s.split(sep).join(""));
    else {
      add(s.replace(sep, "."));
      if (s.split(sep)[1].length === 3) add(s.split(sep).join(""));
    }
  } else add(s);
  return out;
}

export function numbersInText(text) {
  const found = new Set();
  for (const m of text.matchAll(/\d[\d.,'\s  ]*/g)) {
    const parts = m[0].split(/[\s  ]+/).map((p) => p.replace(/^[.,]+|[.,]+$/g, "")).filter(Boolean);
    for (let i = 0; i < parts.length; i++)
      for (let j = i + 1; j <= Math.min(i + 5, parts.length); j++)
        for (const v of parseVariants(parts.slice(i, j).join(""))) found.add(key4(v));
  }
  return found;
}

const MONTHS = [
  "january february march april may june july august september october november december",
  "января февраля марта апреля мая июня июля августа сентября октября ноября декабря",
  "ledna února března dubna května června července srpna září října listopadu prosince",
  "januara februara marta aprila maja juna jula avgusta septembra oktobra novembra decembra",
  "januar februar märz april mai juni juli august september oktober november dezember",
];
const pad = (n) => String(n).padStart(2, "0");

export function dateInText(iso, compactText) {
  const [y, m, d] = iso.split("-").map(Number);
  const v = new Set([`${pad(d)}.${pad(m)}.${y}`, `${d}.${m}.${y}`, `${pad(d)}.${pad(m)}.${String(y).slice(2)}`, iso,
    `${pad(d)}/${pad(m)}/${y}`, `${d}/${m}/${y}`, `${pad(m)}/${pad(d)}/${y}`, `${pad(d)}-${pad(m)}-${y}`]);
  for (const names of MONTHS) {
    const mn = names.split(" ")[m - 1];
    [`${d}${mn}${y}`, `${pad(d)}${mn}${y}`, `${d}.${mn}${y}`, `${mn}${d},${y}`, `${mn}${pad(d)},${y}`].forEach((x) => v.add(x));
  }
  return [...v].some((x) => compactText.includes(x.toLowerCase()));
}

export const compact = (s) => s.replace(/\s+/g, "").toLowerCase();
export const ruDate = (iso) => (iso ? iso.split("-").reverse().join(".") : "");
const isIso = (s) => typeof s === "string" && /^\d{4}-\d\d-\d\d$/.test(s) && !Number.isNaN(Date.parse(s));
const daysBetween = (a, b) => (Date.parse(a) - Date.parse(b)) / 86400000;

// ───────────────────────────── проверки ─────────────────────────────
const LITRE_UNITS = new Set(["л", "l", "lt", "ltr", "litr", "litrů", "litry", "litara", "liter", "liters", "litre", "litres", "литр", "литры", "литров"]);
const TO_LITRES = { "м3": 1000, "м³": 1000, "m3": 1000, "m³": 1000, "куб.м": 1000, "дал": 10, "hl": 100, "ml": 0.001 };

export function litresOf(it) {
  const q = num(it.quantity);
  const unit = (it.unit || "").trim().toLowerCase().replace(/\.$/, "");
  if (q === null) return [null, "количество не указано"];
  if (LITRE_UNITS.has(unit)) return [q, null];
  if (unit in TO_LITRES) return [q * TO_LITRES[unit], `количество ${fmt(q)} ${it.unit} переведено в литры`];
  const vol = num(it.package_volume_liters);
  if (vol) return [q * vol, `литры посчитаны: ${fmt(q)} ${it.unit || "шт"} × ${fmt(vol)} л = ${fmt(q * vol)} л`];
  return [null, `количество в «${it.unit || "?"}», не в литрах`];
}

/** Сумма строки с НДС для колонки «Сумма» + пояснение, если её пришлось вывести. */
export function lineGross(it, inv) {
  const net = num(it.amount_net), vat = num(it.vat_amount), gross = num(it.amount_gross);
  let rate = num(it.vat_rate);
  if (gross !== null || net === null) return [gross, null];
  if (vat !== null) return [r2(net + vat), null];
  const tn = num(inv.total_net), tv = num(inv.total_vat), tg = num(inv.total_gross);
  if (inv.items.length === 1 && tg !== null && (tn === null || close(tn, net, 0.001))) return [tg, [INFO, "сумма с НДС взята из итога счёта"]];
  let note = "";
  if (rate === null && tn && tv !== null) { rate = r2(tv / tn * 100); note = " (ставка из итога счёта)"; }
  if (rate !== null) return [r2(net * (1 + rate / 100)), [INFO, `сумма с НДС посчитана: + ${fmt(rate)}%${note}`]];
  return [net, [WARN, "в «Сумме» — сумма без НДС (НДС в счёте не найден)"]];
}

export function checkInvoice(inv, pdfTextStr, todayIso) {
  const head = [], rows = inv.items.map(() => []);
  const isScan = pdfTextStr.trim().length < 30;
  const nums = isScan ? new Set() : numbersInText(pdfTextStr);
  const ctext = compact(pdfTextStr);
  const inPdf = (v) => isScan || v === null || nums.has(key4(v));

  for (const [k, label] of [["supplier_name", "поставщик"], ["invoice_number", "номер счёта"], ["invoice_date", "дата счёта"], ["currency", "валюта"]])
    if (!inv[k]) head.push([WARN, `не найдено: ${label}`]);
  if (!inv.items.length) head.push([ERROR, "в счёте не найдено ни одной позиции"]);

  const date = (k) => {
    if (!inv[k]) return null;
    if (!isIso(inv[k])) { head.push([ERROR, `непонятная дата: ${inv[k]}`]); return null; }
    return inv[k];
  };
  const dInv = date("invoice_date"), dShip = date("shipment_date");
  if (dShip === null && !inv.shipment_date) head.push([WARN, "дата отгрузки в счёте не найдена"]);
  for (const [d, label] of [[dInv, "дата счёта"], [dShip, "дата отгрузки"]]) {
    if (!d) continue;
    if (daysBetween(d, todayIso) > 1) head.push([ERROR, `${label} ${ruDate(d)} в будущем`]);
    else if (daysBetween(todayIso, d) > 400) head.push([WARN, `${label} ${ruDate(d)} больше года назад`]);
    if (!isScan && !dateInText(d, ctext)) head.push([ERROR, `${label} ${ruDate(d)} не найдена в тексте PDF — проверьте`]);
  }
  if (dInv && dShip && !(daysBetween(dShip, dInv) >= -92 && daysBetween(dShip, dInv) <= 31))
    head.push([WARN, `отгрузка ${ruDate(dShip)} далеко от даты счёта ${ruDate(dInv)}`]);
  if (!isScan && inv.invoice_number && !ctext.includes(compact(inv.invoice_number)))
    head.push([WARN, `номер «${inv.invoice_number}» не найден в тексте PDF`]);

  const totals = { net: 0, vat: 0, gross: 0 }, have = { net: true, vat: true, gross: true };
  inv.items.forEach((it, i) => {
    const q = num(it.quantity), p = num(it.price_per_unit);
    const net = num(it.amount_net), vat = num(it.vat_amount), gross = num(it.amount_gross), rate = num(it.vat_rate);
    const msgs = rows[i];
    const [, note] = litresOf(it);
    if (note) msgs.push([note.includes("не ") ? WARN : INFO, note]);
    if (gross === null && net === null) msgs.push([ERROR, "сумма не найдена"]);
    for (const [v, label] of [[q, "количество"], [num(it.package_volume_liters), "объём упаковки"], [p, "цена"], [net, "сумма без НДС"], [vat, "НДС"], [gross, "сумма"]])
      if (!inPdf(v)) msgs.push([ERROR, `${label} ${fmt(v)} — нет в тексте PDF, возможна ошибка распознавания`]);
    if (q !== null && p !== null) {
      const expected = q * p, tol = 0.01 + q * 0.005;
      let targets = it.price_includes_vat === true ? [gross] : it.price_includes_vat === false ? [net] : [net, gross];
      targets = targets.filter((t) => t !== null);
      if (targets.length && !targets.some((t) => close(expected, t, tol)))
        msgs.push([ERROR, `количество × цена не сходится: ${fmt(q)} × ${fmt(p)} = ${fmt(r2(expected))}, а в счёте ${fmt(targets[0])} (разница ${fmt(r2(Math.abs(targets[0] - expected)))})`]);
    } else if (q !== null && p === null) msgs.push([WARN, "цены нет — арифметику строки проверить нельзя"]);
    if (net !== null && vat !== null && gross !== null && !close(net + vat, gross, 0.02))
      msgs.push([ERROR, `без НДС + НДС ≠ с НДС: ${fmt(net)} + ${fmt(vat)} ≠ ${fmt(gross)}`]);
    if (net !== null && vat !== null && rate !== null && !close(net * rate / 100, vat, 0.02))
      msgs.push([ERROR, `НДС ${fmt(vat)} не равен ${fmt(rate)}% от ${fmt(net)}`]);
    let [g, gNote] = lineGross(it, inv);
    if (gNote && gNote[0] === WARN) g = null;
    for (const [k, v] of [["net", net], ["vat", vat], ["gross", g]]) { if (v === null) have[k] = false; else totals[k] += v; }
  });

  const tol = 0.02 * Math.max(inv.items.length, 1);
  for (const [k, tk, label] of [["net", "total_net", "без НДС"], ["vat", "total_vat", "НДС"], ["gross", "total_gross", "с НДС"]]) {
    const t = num(inv[tk]);
    if (t === null) continue;
    if (!inPdf(t)) head.push([ERROR, `итог ${label} ${fmt(t)} — нет в тексте PDF`]);
    if (have[k] && !close(totals[k], t, tol))
      head.push([ERROR, `сумма строк ${label} ${fmt(r2(totals[k]))} ≠ итогу в счёте ${fmt(t)} (разница ${fmt(r2(Math.abs(t - totals[k])))})`]);
  }
  const tn = num(inv.total_net), tv = num(inv.total_vat), tg = num(inv.total_gross);
  if (tn !== null && tv !== null && tg !== null && !close(tn + tv, tg, 0.02))
    head.push([ERROR, `итог без НДС + НДС ≠ итогу: ${fmt(tn)} + ${fmt(tv)} ≠ ${fmt(tg)}`]);
  const rates = [...new Set(inv.items.map((it) => num(it.vat_rate)).filter((x) => x !== null))];
  if (tn && tv !== null && rates.length === 1 && !close(tn * rates[0] / 100, tv, tol))
    head.push([ERROR, `НДС по счёту ${fmt(tv)} не соответствует ставке от ${fmt(tn)}`]);
  if (tg === null) head.push([WARN, "итоговая сумма счёта не найдена — сверить строки с итогом нельзя"]);
  if (inv.uncertain_fields?.length) head.push([WARN, "модель не уверена в полях: " + inv.uncertain_fields.join(", ")]);
  if (isScan) head.push([INFO, "скан: числа сверены арифметикой (текста в PDF нет)"]);
  return [head, rows];
}

export function worst(levels) {
  let s = OK;
  for (const l of levels) if (RANK[l] > RANK[s]) s = l;
  return s === INFO ? OK : s;
}

/** Позиции счёта → строки таблицы. Значения: числа, даты 'YYYY-MM-DD', текст. */
export function buildRows(inv, head, rows) {
  return inv.items.map((it, i) => {
    const msgs = [...rows[i]];
    const [litres] = litresOf(it);
    const [gross, note] = lineGross(it, inv);
    if (note) msgs.push(note);
    const p = num(it.price_per_unit), q = num(it.quantity);
    const priceL = p !== null && q && litres ? r2(p * q / litres) : null;
    let net = num(it.amount_net);
    if (net === null && gross !== null && num(it.vat_amount) !== null) net = r2(gross - num(it.vat_amount));
    const vat = num(it.vat_amount) ?? (gross !== null && net !== null && gross !== net ? r2(gross - net) : null);
    return {
      values: {
        "Название": it.name, "Сумма": gross, "Литры": litres, "Цена за л": priceL,
        "Поставщик": inv.supplier_name, "№ счёта": inv.invoice_number,
        "Дата": isIso(inv.invoice_date) ? inv.invoice_date : null,
        "Дата отгрузки": isIso(inv.shipment_date) ? inv.shipment_date : null,
        "Валюта": inv.currency,
        // для чужих таблиц — дополнительные поля
        "Сумма без НДС": net, "НДС": vat, "Количество": q, "Ед.": it.unit, "Цена за ед.": p,
      },
      status: worst([...head, ...msgs].map(([l]) => l)),
      messages: msgs,
      edits: [],
    };
  });
}

export function invoiceKey(inv) {
  if (!inv.invoice_number) return null;
  const who = (inv.supplier_tax_id || "").replace(/\D/g, "") || (inv.supplier_name || "").toLowerCase().replace(/[\W_]/gu, "");
  return `${who}|${inv.invoice_number.toLowerCase().replace(/[\W_]/gu, "")}|${(inv.invoice_date || "").slice(0, 4)}`;
}

// ───────────────────────────── ручные правки ─────────────────────────────
export const NUMBER_FIELDS = new Set(["Сумма", "Литры", "Цена за л", "Сумма без НДС", "НДС", "Количество", "Цена за ед."]);
export const DATE_FIELDS = new Set(["Дата", "Дата отгрузки"]);
export const ROW_FIELDS = ["Название", "Сумма", "Литры", "Цена за л"];
export const INVOICE_FIELDS = ["Поставщик", "№ счёта", "Дата", "Дата отгрузки", "Валюта"];

export function parseNumber(s) {
  if (s === null || s === undefined) return null;
  if (typeof s === "number") return s;
  s = String(s).replace(/[\s  ']/g, "");
  if (!s) return null;
  if (s.includes(",") && s.includes(".")) {
    const dec = s.lastIndexOf(",") > s.lastIndexOf(".") ? "," : ".";
    s = s.split(dec === "," ? "." : ",").join("").replace(dec, ".");
  } else s = s.replace(",", ".");
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

export function parseDateInput(s) {
  s = String(s ?? "").trim();
  if (!s) return null;
  if (isIso(s)) return s;
  const m = s.match(/^(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})$/);
  if (!m) return NaN;
  const d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  if (d.getUTCDate() !== +m[1] || d.getUTCMonth() !== +m[2] - 1) return NaN;
  return d.toISOString().slice(0, 10);
}

const show = (field, v) => (v === null || v === undefined || v === "" ? "пусто" : DATE_FIELDS.has(field) ? ruDate(v) : NUMBER_FIELDS.has(field) ? fmt(v) : String(v));

/** Применяет правки {invoice:{поле:значение}, rows:{"0":{поле:значение}}} к копии результата. */
export function applyEdits(res, edits) {
  const out = structuredClone(res);
  const set = (row, field, raw) => {
    let v = NUMBER_FIELDS.has(field) ? parseNumber(raw) : DATE_FIELDS.has(field) ? parseDateInput(raw) : (String(raw ?? "").trim() || null);
    if (Number.isNaN(v)) throw new Error(`${res.filename}: «${raw}» — не ${DATE_FIELDS.has(field) ? "дата" : "число"} (${field})`);
    const old = row.values[field];
    if (v === old || (v !== null && old !== null && NUMBER_FIELDS.has(field) && Math.abs(v - old) < 1e-9)) return;
    row.values[field] = v;
    row.edits.push(`${field}: ${show(field, old)} → ${show(field, v)}`);
    row.status = EDITED;
  };
  for (const [f, v] of Object.entries(edits?.invoice || {})) if (INVOICE_FIELDS.includes(f)) out.rows.forEach((r) => set(r, f, v));
  for (const [i, fields] of Object.entries(edits?.rows || {}))
    for (const [f, v] of Object.entries(fields)) if (ROW_FIELDS.includes(f) && out.rows[+i]) set(out.rows[+i], f, v);
  return out;
}

// ───────────────────────────── разбор одного файла ─────────────────────────────
/**
 * ctx: { key, model, keys: Map, hashes: Map, cacheGet(sha), cachePut(sha, inv), today: 'YYYY-MM-DD' }
 * kind: invoice | dup_file | dup_invoice | not_invoice | error
 */
export async function analyze(bytes, filename, ctx) {
  const sha = await sha256(bytes);
  const res = { id: `${sha.slice(0, 12)}-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, sha, filename,
    kind: "invoice", message: "", key: null, invoice: {}, head: [], rows: [], status: OK, usage: {}, include: true };
  if (ctx.hashes.has(sha)) {
    const where = ctx.hashes.get(sha);
    return { ...res, kind: "dup_file", include: false, message: `этот же файл уже ${where}` };
  }
  let inv = await ctx.cacheGet(sha);
  if (inv) res.usage = { cached: true };
  else {
    try {
      const r = await extract(bytes, filename, ctx.key, ctx.model);
      inv = r.inv; res.usage = r.usage;
      await ctx.cachePut(sha, inv);
    } catch (e) {
      return { ...res, kind: "error", include: false, message: `сбой распознавания: ${e.message}` };
    }
  }
  res.invoice = { supplier: inv.supplier_name, number: inv.invoice_number, date: inv.invoice_date, shipment_date: inv.shipment_date,
    shipment_label: inv.shipment_date_label, currency: inv.currency, total: inv.total_gross, notes: inv.notes };
  if (!inv.is_invoice) return { ...res, kind: "not_invoice", include: false, message: "не похоже на счёт" };
  const key = invoiceKey(inv);
  res.key = key;
  if (key && ctx.keys.has(key))
    return { ...res, kind: "dup_invoice", include: false, message: `счёт № ${inv.invoice_number} уже ${ctx.keys.get(key)}` };
  const [head, rows] = checkInvoice(inv, await pdfText(bytes), ctx.today);
  res.head = head;
  res.rows = buildRows(inv, head, rows);
  res.status = worst(res.rows.map((r) => r.status).concat(res.rows.length ? [] : [ERROR]));
  return res;
}
