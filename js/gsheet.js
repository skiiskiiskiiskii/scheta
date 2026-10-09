// Запись в Google Таблицу через маленький скрипт (Google Apps Script) внутри самой таблицы.
// Сайт отправляет строки скрипту, скрипт дописывает их в конец листа с оформлением последней строки.
// Скрипт проверяет секретный ключ, поэтому писать в таблицу может только этот сайт с этим ключом.

export function newToken() {
  const a = new Uint8Array(18);
  crypto.getRandomValues(a);
  return [...a].map((b) => b.toString(16).padStart(2, "0")).join("");
}

export function scriptCode(token) {
  return `// «Счета → Excel»: приём строк с сайта. Вставьте этот код целиком, ничего не меняя.
const TOKEN = "${token}";

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: "плохой запрос" }); }
  if (req.token !== TOKEN) return out_({ error: "неверный ключ подключения" });
  try {
    if (req.action === "info") return out_(info_(req));
    if (req.action === "append") return out_(append_(req));
    return out_({ error: "неизвестное действие" });
  } catch (err) {
    return out_({ error: String((err && err.message) || err) });
  }
}

function doGet() { return out_({ ok: true, hint: "Счета → Excel: подключение работает" }); }

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function info_(req) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var all = ss.getSheets();
  var sheets = all.map(function (s) { return { name: s.getName(), hidden: s.isSheetHidden() }; });
  var sh = req.sheet ? ss.getSheetByName(req.sheet) : all.filter(function (s) { return !s.isSheetHidden(); })[0];
  var values = [];
  if (sh && sh.getLastRow() > 0 && sh.getLastColumn() > 0) {
    values = sh.getRange(1, 1, Math.min(sh.getLastRow(), 5000), Math.min(sh.getLastColumn(), 40)).getDisplayValues();
  }
  return { name: ss.getName(), url: ss.getUrl(), sheets: sheets, sheet: sh ? sh.getName() : null, values: values };
}

function append_(req) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(req.sheet);
    if (!sh) throw new Error("лист «" + req.sheet + "» не найден");
    var cols = req.columns, header = req.headerRow || 1, maxRows = sh.getMaxRows();
    // последняя заполненная строка — смотрим только колонки, в которые пишем
    var last = header;
    cols.forEach(function (c) {
      var v = sh.getRange(1, c, maxRows, 1).getDisplayValues();
      for (var r = v.length; r > last; r--) { if (v[r - 1][0] !== "") { last = r; break; } }
    });
    var first = last + 1, n = req.rows.length;
    if (first + n - 1 > maxRows) sh.insertRowsAfter(maxRows, first + n - 1 - maxRows);
    var width = Math.max(sh.getLastColumn(), Math.max.apply(null, cols));
    if (last > header) sh.getRange(last, 1, 1, width).copyFormatToRange(sh, 1, width, first, first + n - 1);
    req.rows.forEach(function (row, i) {
      cols.forEach(function (c, j) {
        var v = row[j], type = req.types[j];
        if (v === null || v === undefined || v === "") return;
        var cell = sh.getRange(first + i, c);
        if (type === "date") {
          var p = String(v).split("-");
          cell.setValue((Date.UTC(+p[0], +p[1] - 1, +p[2]) - Date.UTC(1899, 11, 30)) / 86400000);
          if (!/[dmy]/i.test(cell.getNumberFormat())) cell.setNumberFormat("dd.mm.yyyy");
        } else if (type === "number") {
          cell.setValue(Number(v));
        } else {
          v = String(v);
          cell.setValue(/^[=+\\-@]/.test(v) ? "'" + v : v);
        }
      });
    });
    SpreadsheetApp.flush();
    return { first: first, last: first + n - 1, url: ss.getUrl() + "#gid=" + sh.getSheetId() };
  } finally {
    lock.releaseLock();
  }
}
`;
}

/** Запрос к скрипту таблицы. Простой POST text/plain — Google отвечает с разрешением для браузера. */
export async function call(url, token, payload) {
  let r;
  try {
    r = await fetch(url, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ token, ...payload }), redirect: "follow" });
  } catch {
    throw new Error("не удалось связаться с Google Таблицей — проверьте ссылку и интернет. Если ссылка верная, проверьте, что при публикации выбран доступ «Все» (Anyone)");
  }
  const text = await r.text();
  let data;
  try { data = JSON.parse(text); } catch {
    throw new Error(/<html/i.test(text) ? "по ссылке открылась страница Google, а не скрипт — проверьте, что вставлена ссылка «Веб-приложение» (оканчивается на /exec) и доступ «Все» (Anyone)" : "непонятный ответ от таблицы");
  }
  if (data.error) throw new Error(data.error);
  return data;
}

export const isScriptUrl = (u) => /^https:\/\/script\.google(usercontent)?\.com\/.+/.test((u || "").trim());
