/**
 * Google Apps Script для веб-приложения «Раздельный учёт».
 * Лист: первая строка — заголовки (см. SHEET_COLUMNS).
 *
 * doGet:  ?action=getRows  → JSON-массив объектов со всеми полями.
 * doPost: тело JSON { "action": "upsert", "rows": [ {...}, ... ] }
 *         upsert по паре period + order (обновление строки или append).
 */

var SHEET_COLUMNS = [
  'period',
  'order',
  'shipment',
  'materials',
  'processing',
  'other',
  'salary',
  'account25',
  'account26',
  'account44',
  'total_cost',
  'writeoff',
  'profit',
  'profit_percent',
  'status'
];

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function ensureHeaders(sheet) {
  var n = SHEET_COLUMNS.length;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, n).setValues([SHEET_COLUMNS]);
    return;
  }
  var first = sheet.getRange(1, 1, 1, n).getValues()[0];
  var ok = true;
  for (var i = 0; i < n; i++) {
    if (String(first[i] || '')
      .trim()
      .toLowerCase() !== SHEET_COLUMNS[i]) {
      ok = false;
      break;
    }
  }
  if (!ok) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, n).setValues([SHEET_COLUMNS]);
  }
}

function headerIndex(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), SHEET_COLUMNS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '')
      .trim()
      .toLowerCase();
    if (h) map[h] = c;
  }
  return map;
}

function rowToObject(values, idx) {
  var o = {};
  for (var k = 0; k < SHEET_COLUMNS.length; k++) {
    var name = SHEET_COLUMNS[k];
    var col = idx[name];
    o[name] = col === undefined ? '' : values[col];
  }
  return o;
}

function normalizePeriodCell(p) {
  if (Object.prototype.toString.call(p) === '[object Date]' && !isNaN(p.getTime())) {
    return Utilities.formatDate(p, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var s = String(p == null ? '' : p)
    .replace(/\u00a0/g, ' ')
    .trim();
  var m = s.match(/^(\d{4})-(\d{2})/);
  if (m) return m[1] + '-' + m[2];
  return s;
}

function doGet(e) {
  e = e || { parameter: {} };
  if (e.parameter.action !== 'getRows') {
    return jsonOut({ ok: true });
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureHeaders(sheet);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return jsonOut([]);
  }
  var idx = headerIndex(sheet);
  var out = [];
  for (var r = 1; r < data.length; r++) {
    out.push(rowToObject(data[r], idx));
  }
  return jsonOut(out);
}

function sheetRowArrayFromPayload(row) {
  var arr = [];
  for (var k = 0; k < SHEET_COLUMNS.length; k++) {
    var name = SHEET_COLUMNS[k];
    arr.push(row[name] != null && row[name] !== '' ? row[name] : '');
  }
  return arr;
}

function upsertRows(sheet, rows) {
  ensureHeaders(sheet);
  var data = sheet.getDataRange().getValues();
  var idx = headerIndex(sheet);
  var pCol = idx['period'];
  var oCol = idx['order'];
  if (pCol === undefined || oCol === undefined) {
    return jsonOut({ error: 'missing period or order column' });
  }
  var n = SHEET_COLUMNS.length;
  var updated = 0;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var period = normalizePeriodCell(row.period);
    var order = String(row.order == null ? '' : row.order).trim();
    if (!order) continue;
    var rowArr = sheetRowArrayFromPayload(row);
    var found = -1;
    for (var r = 1; r < data.length; r++) {
      var pv = normalizePeriodCell(data[r][pCol]);
      var ov = String(data[r][oCol] == null ? '' : data[r][oCol]).trim();
      if (pv === period && ov === order) {
        found = r + 1;
        break;
      }
    }
    if (found > 0) {
      sheet.getRange(found, 1, found, n).setValues([rowArr]);
      for (var c = 0; c < n; c++) {
        data[found - 1][c] = rowArr[c];
      }
      updated++;
    } else {
      sheet.appendRow(rowArr);
      var newRow = [];
      for (var z = 0; z < n; z++) newRow.push(rowArr[z]);
      data.push(newRow);
    }
  }
  return jsonOut({ ok: true, rows: rows.length, updated: updated });
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ error: 'invalid json' });
  }
  if (body && body.action === 'upsert' && Array.isArray(body.rows)) {
    return upsertRows(sheet, body.rows);
  }
  if (Array.isArray(body)) {
    return upsertRows(sheet, body);
  }
  if (body && typeof body === 'object' && body.period != null && body.order != null) {
    return upsertRows(sheet, [body]);
  }
  return jsonOut({ error: 'expected { action: "upsert", rows: [...] }' });
}
