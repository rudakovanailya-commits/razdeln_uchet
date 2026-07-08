/**
 * Google Apps Script для веб-приложения «Раздельный учёт».
 *
 * Script Properties (Project Settings → Script properties):
 *   SPREADSHEET_ID  — ID Google Таблицы (обязательно для Web App)
 *   DATA_SHEET_NAME — имя листа с данными (необязательно; иначе первый лист)
 *   ADMIN_DELETE_TOKEN — токен удаления периода
 *
 * doGet:
 *   ?action=getRows
 *   ?action=getPeriodData&period=YYYY-MM
 *   ?action=getPeriodDataList
 *   ?action=getActionLog&limit=100&period=YYYY-MM&user_name=...&action=...
 *   ?action=debugLastError
 *   ?action=getLog&limit=50
 *
 * doPost:
 *   { "action": "replacePeriod", "period": "YYYY-MM", "rows": [...] }
 *   { "action": "savePeriodData", "period": "YYYY-MM", "periodData": { ...полный объект периода v2... } }
 *   (legacy: periodData: { "json": "...", "updated_by": "v2", "format_version": 2 })
 *   { "action": "appendActionLog", "logEntry": { ... } }
 *   { "action": "upsert", "rows": [...] }
 *   { "action": "deletePeriod", "period": "YYYY-MM", "token": "..." }
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

var DEBUG_LOG_SHEET_NAME = '_log';
var DEBUG_LOG_HEADERS = [
  'timestamp',
  'action',
  'period',
  'rowsCount',
  'validRowsCount',
  'error',
  'stack',
  'message',
  'data'
];

var DEBUG_LAST_ERROR_KEY = 'DEBUG_LAST_ERROR';

var PERIOD_DATA_SHEET_NAME = 'period_data';
var PERIOD_DATA_HEADERS = [
  'period',
  'json',
  'updated_at',
  'updated_by',
  'format_version'
];

var ACTION_LOG_SHEET_NAME = '_action_log';
var ACTION_LOG_HEADERS = [
  'timestamp',
  'period',
  'action',
  'status',
  'user_name',
  'device_id',
  'source',
  'details',
  'rows_count',
  'browser',
  'timezone',
  'app_version'
];

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function getSpreadsheetConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    id: props.getProperty('SPREADSHEET_ID') || props.getProperty('SHEET_ID') || '',
    sheetName: props.getProperty('DATA_SHEET_NAME') || props.getProperty('SHEET_NAME') || ''
  };
}

/** Web App: openById. Редактор: fallback на active spreadsheet. */
function getDataSpreadsheet() {
  var cfg = getSpreadsheetConfig();
  if (cfg.id) {
    return SpreadsheetApp.openById(cfg.id);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getDataSheet() {
  var ss = getDataSpreadsheet();
  var cfg = getSpreadsheetConfig();
  if (cfg.sheetName) {
    var named = ss.getSheetByName(cfg.sheetName);
    if (!named) {
      throw new Error('DATA_SHEET_NAME not found: ' + cfg.sheetName);
    }
    return named;
  }
  var sheets = ss.getSheets();
  if (!sheets || !sheets.length) {
    throw new Error('Spreadsheet has no sheets');
  }
  return sheets[0];
}

function getPeriodDataSheet() {
  var ss = getDataSpreadsheet();
  var sh = ss.getSheetByName(PERIOD_DATA_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PERIOD_DATA_SHEET_NAME);
  }
  ensurePeriodDataHeaders(sh);
  return sh;
}

function ensurePeriodDataHeaders(sheet) {
  if (!sheet) return;
  var n = PERIOD_DATA_HEADERS.length;
  var lastCol = Math.max(sheet.getLastColumn(), n);
  var headerRow =
    lastCol >= n
      ? sheet.getRange(1, 1, 1, n).getValues()[0]
      : [];
  var needsHeader = sheet.getLastRow() < 1;
  if (!needsHeader && headerRow.length >= n) {
    for (var i = 0; i < n; i++) {
      if (
        String(headerRow[i] || '')
          .trim()
          .toLowerCase() !== PERIOD_DATA_HEADERS[i]
      ) {
        needsHeader = true;
        break;
      }
    }
  }
  if (needsHeader) {
    sheet.getRange(1, 1, 1, n).setValues([PERIOD_DATA_HEADERS]);
  }
  ensurePeriodColumnPlainText(sheet, 1);
}

/**
 * period всегда текст YYYY-MM (не дата Sheets).
 * Date/ISO → через Session.getScriptTimeZone(), чтобы не сдвигать месяц.
 */
function normalizePeriodText(value) {
  if (value == null || value === '') return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    try {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
    } catch (e) {
      return '';
    }
  }

  var s = String(value)
    .replace(/\u00a0/g, ' ')
    .replace(/^\uFEFF/, '')
    .trim();
  if (!s) return '';
  if (s.charAt(0) === "'") s = s.substring(1).trim();
  if (!s) return '';

  // Уже чистый период
  if (/^\d{4}-\d{2}$/.test(s)) return s;

  // ISO / datetime от Sheets (напр. 2025-11-30T21:00:00.000Z = 01.12.2025 локально)
  if (/^\d{4}-\d{2}-\d{2}T/.test(s) || /Z$/i.test(s) || /\+\d{2}:?\d{2}$/.test(s)) {
    var dIso = new Date(s);
    if (!isNaN(dIso.getTime())) {
      try {
        return Utilities.formatDate(dIso, Session.getScriptTimeZone(), 'yyyy-MM');
      } catch (e2) {
        /* fall through */
      }
    }
  }

  // YYYY-MM-DD без времени — берём год-месяц как в дате таблицы
  var ymd = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (ymd) return ymd[1] + '-' + ymd[2];

  // YYYY-MM... prefix (не datetime)
  var ym = s.match(/^(\d{4})-(\d{2})(?!\d)/);
  if (ym) return ym[1] + '-' + ym[2];

  // DD.MM.YYYY или DD/MM/YYYY
  var dmy = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})/);
  if (dmy) {
    return dmy[3] + '-' + ('0' + dmy[2]).slice(-2);
  }

  // MM.YYYY
  var my = s.match(/^(\d{1,2})[./](\d{4})$/);
  if (my) {
    return my[2] + '-' + ('0' + my[1]).slice(-2);
  }

  return '';
}

/** Алиас: все вызовы нормализации периода идут через normalizePeriodText. */
function normalizePeriodCell(p) {
  return normalizePeriodText(p);
}

function normalizeLogPeriod(value) {
  return normalizePeriodText(value);
}

function normalizePeriodDataKey(periodParam) {
  var p = normalizePeriodText(periodParam);
  if (!/^\d{4}-\d{2}$/.test(p)) return '';
  return p;
}

/** Колонка period — plain text (@), чтобы Sheets не превращал 2025-12 в дату. */
function ensurePeriodColumnPlainText(sheet, columnIndex) {
  if (!sheet || !columnIndex || columnIndex < 1) return;
  try {
    sheet.getRange(1, columnIndex, sheet.getMaxRows(), 1).setNumberFormat('@');
  } catch (e) {
    try {
      sheet.getRange(1, columnIndex, Math.max(sheet.getLastRow(), 1), 1).setNumberFormat('@');
    } catch (e2) {
      /* ignore */
    }
  }
}

function periodTextForSheetWrite(value) {
  return normalizePeriodText(value);
}

/** Возвращает объект строки period_data или null, если период не найден. */
function getPeriodDataRow(periodParam) {
  var period = normalizePeriodDataKey(periodParam);
  if (!period) return null;

  var sheet = getPeriodDataSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var n = PERIOD_DATA_HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow, n).getValues();
  var latest = null;
  var latestAt = '';

  for (var r = 0; r < values.length; r++) {
    var rowPeriod = normalizePeriodText(values[r][0]);
    if (rowPeriod !== period) continue;
    var json = values[r][1] != null ? String(values[r][1]) : '';
    if (!json) continue;
    var updatedAt = values[r][2] != null ? String(values[r][2]) : '';
    var candidate = {
      period: rowPeriod,
      json: json,
      updated_at: updatedAt,
      updated_by: values[r][3] != null ? String(values[r][3]) : '',
      format_version:
        values[r][4] != null && values[r][4] !== '' ? String(values[r][4]) : ''
    };
    if (!latest || String(updatedAt) >= String(latestAt)) {
      latest = candidate;
      latestAt = updatedAt;
    }
  }

  return latest;
}

/** Список периодов из period_data (последняя версия по updated_at на период). */
function getPeriodDataList() {
  var sheet = getPeriodDataSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: true, periods: [] };
  }

  var n = PERIOD_DATA_HEADERS.length;
  var values = sheet.getRange(2, 1, lastRow, n).getValues();
  var byPeriod = {};

  for (var r = 0; r < values.length; r++) {
    var rowPeriod = normalizePeriodText(values[r][0]);
    if (!/^\d{4}-\d{2}$/.test(rowPeriod)) continue;
    var json = values[r][1] != null ? String(values[r][1]) : '';
    if (!json) continue;
    var updatedAt = values[r][2] != null ? String(values[r][2]) : '';
    var candidate = {
      period: rowPeriod,
      updated_at: updatedAt,
      updated_by: values[r][3] != null ? String(values[r][3]) : '',
      format_version:
        values[r][4] != null && values[r][4] !== '' ? String(values[r][4]) : ''
    };
    var prev = byPeriod[rowPeriod];
    if (!prev || String(updatedAt) >= String(prev.updated_at)) {
      byPeriod[rowPeriod] = candidate;
    }
  }

  var periods = Object.keys(byPeriod)
    .sort()
    .map(function (k) {
      return byPeriod[k];
    });

  return { ok: true, periods: periods };
}

function getActionLogSheet() {
  var ss = getDataSpreadsheet();
  var sh = ss.getSheetByName(ACTION_LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ACTION_LOG_SHEET_NAME);
  }
  ensureActionLogHeaders(sh);
  return sh;
}

function ensureActionLogHeaders(sheet) {
  if (!sheet) return;
  var n = ACTION_LOG_HEADERS.length;
  var lastCol = Math.max(sheet.getLastColumn(), n);
  var headerRow =
    lastCol >= n
      ? sheet.getRange(1, 1, 1, n).getValues()[0]
      : [];
  var needsHeader = sheet.getLastRow() < 1;
  if (!needsHeader && headerRow.length >= n) {
    for (var i = 0; i < n; i++) {
      if (
        String(headerRow[i] || '')
          .trim()
          .toLowerCase() !== ACTION_LOG_HEADERS[i]
      ) {
        needsHeader = true;
        break;
      }
    }
  }
  if (needsHeader) {
    sheet.getRange(1, 1, 1, n).setValues([ACTION_LOG_HEADERS]);
  }
  // period — колонка B (индекс 2)
  ensurePeriodColumnPlainText(sheet, 2);
}

function appendActionLogEntry(logEntry) {
  var e = logEntry || {};
  var sheet = getActionLogSheet();
  var periodText = periodTextForSheetWrite(e.period);
  var row = [
    e.timestamp != null ? String(e.timestamp) : new Date().toISOString(),
    periodText,
    e.action != null ? String(e.action) : '',
    e.status != null ? String(e.status) : '',
    e.user_name != null ? String(e.user_name) : '',
    e.device_id != null ? String(e.device_id) : '',
    e.source != null ? String(e.source) : '',
    e.details != null ? String(e.details) : '',
    e.rows_count != null && e.rows_count !== '' ? String(e.rows_count) : '',
    e.browser != null ? String(e.browser) : '',
    e.timezone != null ? String(e.timezone) : '',
    e.app_version != null ? String(e.app_version) : ''
  ];
  ensurePeriodColumnPlainText(sheet, 2);
  sheet.appendRow(row);
  var last = sheet.getLastRow();
  if (last >= 1) {
    sheet.getRange(last, 2).setNumberFormat('@').setValue(periodText);
  }
  return jsonOut({ ok: true });
}

/**
 * Читает лист _action_log (без создания, если листа нет).
 * Параметры: limit (по умолчанию 100), period, user_name, action.
 * Новые записи первыми.
 */
function getActionLog(params) {
  params = params || {};
  var limit = Number(params.limit);
  if (!isFinite(limit) || limit <= 0) limit = 100;
  if (limit > 500) limit = 500;

  var filterPeriod = normalizePeriodText(params.period);
  var filterUser = String(params.user_name || '').trim().toLowerCase();
  var filterAction = String(params.action || '').trim().toLowerCase();

  var ss = getDataSpreadsheet();
  var sheet = ss.getSheetByName(ACTION_LOG_SHEET_NAME);
  if (!sheet) {
    return { ok: true, logs: [] };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: true, logs: [] };
  }

  var n = ACTION_LOG_HEADERS.length;
  var lastCol = Math.max(sheet.getLastColumn(), n);
  var values = sheet.getRange(2, 1, lastRow, Math.min(lastCol, n)).getValues();
  var logs = [];

  function actionLogCellToString(v) {
    if (v == null || v === '') return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      try {
        return v.toISOString();
      } catch (dateErr) {
        return String(v);
      }
    }
    return String(v);
  }

  function isEmptyActionLogEntry(entry) {
    if (!entry) return true;
    var keys = ['timestamp', 'period', 'action', 'status', 'user_name', 'details'];
    for (var i = 0; i < keys.length; i++) {
      var v = String(entry[keys[i]] == null ? '' : entry[keys[i]]).trim();
      if (v && v !== '—' && v !== '-') return false;
    }
    return true;
  }

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var entry = {
      timestamp: actionLogCellToString(row[0]),
      period: normalizePeriodText(row[1]),
      action: actionLogCellToString(row[2]),
      status: actionLogCellToString(row[3]),
      user_name: actionLogCellToString(row[4]),
      device_id: actionLogCellToString(row[5]),
      source: actionLogCellToString(row[6]),
      details: actionLogCellToString(row[7]),
      rows_count: actionLogCellToString(row[8]),
      browser: actionLogCellToString(row[9]),
      timezone: actionLogCellToString(row[10]),
      app_version: actionLogCellToString(row[11])
    };

    if (isEmptyActionLogEntry(entry)) continue;
    if (filterPeriod && entry.period !== filterPeriod) continue;
    if (filterUser && String(entry.user_name).toLowerCase() !== filterUser) continue;
    if (filterAction && String(entry.action).toLowerCase() !== filterAction) continue;

    logs.push(entry);
    if (logs.length >= limit) break;
  }

  return { ok: true, logs: logs };
}

/** Извлекает json и метаданные из periodData (полный объект v2 или legacy-обёртка). */
function parseSavePeriodDataPayload(periodData) {
  if (periodData == null) {
    return { ok: false, error: 'periodData_required' };
  }

  var updatedBy = 'v2';
  var formatVersion = '2';
  var jsonStr = '';

  if (typeof periodData === 'string') {
    jsonStr = periodData;
  } else if (typeof periodData === 'object' && !Array.isArray(periodData)) {
    if (
      periodData.json != null &&
      typeof periodData.json === 'string' &&
      !periodData.rows &&
      !periodData.globals
    ) {
      jsonStr = String(periodData.json);
      if (periodData.updated_by != null && String(periodData.updated_by).trim() !== '') {
        updatedBy = String(periodData.updated_by).trim();
      }
      if (
        periodData.format_version != null &&
        String(periodData.format_version).trim() !== ''
      ) {
        formatVersion = String(periodData.format_version).trim();
      }
    } else if (periodData.rows != null || periodData.globals != null) {
      if (periodData.meta && periodData.meta.updatedBy) {
        updatedBy = String(periodData.meta.updatedBy);
      } else if (periodData.updated_by) {
        updatedBy = String(periodData.updated_by);
      }
      if (periodData.format_version != null) {
        formatVersion = String(periodData.format_version);
      }
      jsonStr = JSON.stringify(periodData);
    } else {
      jsonStr = JSON.stringify(periodData);
    }
  } else {
    return { ok: false, error: 'invalid_periodData' };
  }

  if (!jsonStr || String(jsonStr).trim() === '') {
    return { ok: false, error: 'json_required' };
  }

  return {
    ok: true,
    jsonStr: String(jsonStr),
    updatedBy: updatedBy,
    formatVersion: formatVersion
  };
}

function upsertPeriodData(periodParam, jsonStr, updatedBy, formatVersion) {
  var period = normalizePeriodDataKey(periodParam);
  if (!period) {
    return jsonOut({ ok: false, error: 'invalid_period' });
  }
  if (jsonStr == null || String(jsonStr).trim() === '') {
    return jsonOut({ ok: false, error: 'json_required' });
  }

  var sheet = getPeriodDataSheet();
  ensurePeriodDataHeaders(sheet);
  var updatedAt = new Date().toISOString();
  var by =
    updatedBy != null && String(updatedBy).trim() !== ''
      ? String(updatedBy).trim()
      : 'v2';
  var fv =
    formatVersion != null && String(formatVersion).trim() !== ''
      ? String(formatVersion).trim()
      : '2';
  var json = String(jsonStr);

  var lastRow = sheet.getLastRow();
  var targetRow = -1;

  if (lastRow >= 2) {
    var periodCol = sheet.getRange(2, 1, lastRow, 1).getValues();
    for (var r = 0; r < periodCol.length; r++) {
      if (normalizePeriodText(periodCol[r][0]) === period) {
        targetRow = r + 2;
        break;
      }
    }
  }

  ensurePeriodColumnPlainText(sheet, 1);
  var rowValues = [period, json, updatedAt, by, fv];
  if (targetRow > 0) {
    sheet
      .getRange(targetRow, 1, 1, rowValues.length)
      .setValues([rowValues]);
    sheet.getRange(targetRow, 1).setNumberFormat('@').setValue(period);
  } else {
    sheet.appendRow(rowValues);
    var appendedRow = sheet.getLastRow();
    if (appendedRow >= 1) {
      sheet.getRange(appendedRow, 1).setNumberFormat('@').setValue(period);
    }
  }

  writeDebugLog('savePeriodData: upserted', {
    action: 'savePeriodData',
    period: period,
    data: {
      updated_at: updatedAt,
      updated_by: by,
      format_version: fv,
      jsonLength: json.length,
      row: targetRow > 0 ? targetRow : lastRow + 1
    }
  });

  return jsonOut({
    ok: true,
    period: period,
    updated_at: updatedAt,
    updated_by: by,
    format_version: fv
  });
}

function setLastError(obj) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      DEBUG_LAST_ERROR_KEY,
      JSON.stringify(obj)
    );
  } catch (e) {
    Logger.log('setLastError failed: ' + e);
  }
}

function getLastErrorObject() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(DEBUG_LAST_ERROR_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return { error: 'failed_to_read_DEBUG_LAST_ERROR', detail: String(e) };
  }
}

function getOrCreateLogSheet() {
  var ss = getDataSpreadsheet();
  var sh = ss.getSheetByName(DEBUG_LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(DEBUG_LOG_SHEET_NAME);
    sh.getRange(1, 1, 1, DEBUG_LOG_HEADERS.length).setValues([DEBUG_LOG_HEADERS]);
  }
  return sh;
}

function writeDebugLog(message, data) {
  data = data || {};
  var logRow = {
    timestamp: new Date().toISOString(),
    action: data.action || '',
    period: data.period != null ? String(data.period) : '',
    rowsCount: data.rowsCount != null ? data.rowsCount : '',
    validRowsCount: data.validRowsCount != null ? data.validRowsCount : '',
    error: data.error != null ? String(data.error) : '',
    stack: data.stack != null ? String(data.stack) : '',
    message: message != null ? String(message) : '',
    data: data.data ? JSON.stringify(data.data) : ''
  };

  Logger.log(logRow);

  try {
    var sh = getOrCreateLogSheet();
    sh.appendRow([
      logRow.timestamp,
      logRow.action,
      logRow.period,
      logRow.rowsCount,
      logRow.validRowsCount,
      logRow.error,
      logRow.stack,
      logRow.message,
      logRow.data
    ]);
  } catch (logErr) {
    Logger.log('writeDebugLog sheet failed: ' + logErr);
  }

  if (data.error) {
    setLastError({
      timestamp: logRow.timestamp,
      message: logRow.message,
      error: logRow.error,
      stack: logRow.stack,
      action: logRow.action,
      period: logRow.period,
      data: data.data || null
    });
  }
}

function getLogEntries(limit) {
  limit = limit || 50;
  try {
    var ss = getDataSpreadsheet();
    var sh = ss.getSheetByName(DEBUG_LOG_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) {
      return { ok: true, entries: [] };
    }
    var lastRow = sh.getLastRow();
    var startRow = Math.max(2, lastRow - limit + 1);
    var values = sh.getRange(startRow, 1, lastRow, DEBUG_LOG_HEADERS.length).getValues();
    var entries = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      entries.push({
        timestamp: row[0],
        action: row[1],
        period: row[2],
        rowsCount: row[3],
        validRowsCount: row[4],
        error: row[5],
        stack: row[6],
        message: row[7],
        data: row[8]
      });
    }
    return { ok: true, entries: entries };
  } catch (err) {
    return { ok: false, error: String(err.message || err), stack: String(err.stack || '') };
  }
}

function getAdminDeleteToken() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_DELETE_TOKEN') || '';
}

function isValidDeleteToken(token) {
  var expected = getAdminDeleteToken();
  if (!expected) return false;
  return String(token == null ? '' : token) === expected;
}

function ensureSheetColumnCount(sheet, columnCount) {
  var need = columnCount || SHEET_COLUMNS.length;
  if (sheet.getMaxColumns() < need) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
  }
}

function ensureHeaders(sheet) {
  var n = SHEET_COLUMNS.length;
  ensureSheetColumnCount(sheet, n);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, n).setValues([SHEET_COLUMNS]);
    ensurePeriodColumnPlainText(sheet, 1);
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
  ensurePeriodColumnPlainText(sheet, 1);
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
    var raw = col === undefined ? '' : values[col];
    o[name] = name === 'period' ? normalizePeriodText(raw) : raw;
  }
  return o;
}

function periodHasFinalStatusInSheet(data, idx, pCol, sCol, period) {
  if (sCol === undefined) return false;
  for (var r = 1; r < data.length; r++) {
    if (normalizePeriodCell(data[r][pCol]) !== period) continue;
    var st = String(data[r][sCol] == null ? '' : data[r][sCol])
      .trim()
      .toLowerCase();
    if (st === 'final') return true;
  }
  return false;
}

function deletePeriodRowsAuthorized(sheet, periodParam) {
  ensureHeaders(sheet);
  var period = normalizePeriodCell(periodParam);
  if (!period) {
    return jsonOut({ ok: false, error: 'period_required', deleted: 0 });
  }
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return jsonOut({ ok: true, deleted: 0 });
  }
  var idx = headerIndex(sheet);
  var pCol = idx['period'];
  var sCol = idx['status'];
  if (pCol === undefined) {
    return jsonOut({ ok: false, error: 'missing period column', deleted: 0 });
  }
  if (periodHasFinalStatusInSheet(data, idx, pCol, sCol, period)) {
    return jsonOut({ ok: false, error: 'period_final', deleted: 0 });
  }
  var toDelete = [];
  for (var r = 1; r < data.length; r++) {
    if (normalizePeriodCell(data[r][pCol]) === period) {
      toDelete.push(r + 1);
    }
  }
  if (!toDelete.length) {
    return jsonOut({ ok: true, deleted: 0 });
  }
  toDelete.sort(function(a, b) {
    return b - a;
  });
  for (var i = 0; i < toDelete.length; i++) {
    sheet.deleteRow(toDelete[i]);
  }
  return jsonOut({ ok: true, deleted: toDelete.length });
}

function deletePeriodPost(sheet, body) {
  body = body || {};
  if (!isValidDeleteToken(body.token)) {
    return jsonOut({ ok: false, error: 'forbidden', deleted: 0 });
  }
  return deletePeriodRowsAuthorized(sheet, body.period);
}

function doGet(e) {
  e = e || { parameter: {} };
  var action = String(e.parameter.action || '');

  if (action === 'deletePeriod') {
    return jsonOut({ ok: false, error: 'delete_not_allowed_via_get', deleted: 0 });
  }

  if (action === 'debugLastError') {
    return jsonOut({ ok: true, lastError: getLastErrorObject() });
  }

  if (action === 'getLog') {
    var limit = Number(e.parameter.limit) || 50;
    return jsonOut(getLogEntries(limit));
  }

  if (action === 'getPeriodData') {
    var periodParam = String(e.parameter.period || '').trim();
    if (!normalizePeriodDataKey(periodParam)) {
      return jsonOut({ ok: false, error: 'invalid_period' });
    }
    try {
      var periodRow = getPeriodDataRow(periodParam);
      if (!periodRow) {
        return jsonOut({ ok: false, error: 'not_found' });
      }
      return jsonOut({
        ok: true,
        period: periodRow.period,
        json: periodRow.json,
        updated_at: periodRow.updated_at,
        updated_by: periodRow.updated_by,
        format_version: periodRow.format_version
      });
    } catch (err) {
      var pdMsg = String(err.message || err);
      var pdStack = String(err.stack || '');
      Logger.log(pdStack || pdMsg);
      writeDebugLog('doGet getPeriodData failed', {
        action: 'doGet',
        period: periodParam,
        error: pdMsg,
        stack: pdStack
      });
      return jsonOut({ ok: false, error: pdMsg, stack: pdStack });
    }
  }

  if (action === 'getPeriodDataList') {
    try {
      return jsonOut(getPeriodDataList());
    } catch (err) {
      var pdlMsg = String(err.message || err);
      var pdlStack = String(err.stack || '');
      Logger.log(pdlStack || pdlMsg);
      writeDebugLog('doGet getPeriodDataList failed', {
        action: 'doGet',
        error: pdlMsg,
        stack: pdlStack
      });
      return jsonOut({ ok: false, error: pdlMsg, stack: pdlStack });
    }
  }

  if (action === 'getActionLog') {
    try {
      return jsonOut(
        getActionLog({
          limit: e.parameter.limit,
          period: e.parameter.period,
          user_name: e.parameter.user_name,
          action: e.parameter.actionFilter || e.parameter.log_action || e.parameter.filter_action
        })
      );
    } catch (err) {
      var alMsg = String(err.message || err);
      var alStack = String(err.stack || '');
      Logger.log(alStack || alMsg);
      writeDebugLog('doGet getActionLog failed', {
        action: 'doGet',
        error: alMsg,
        stack: alStack
      });
      return jsonOut({ ok: false, error: alMsg, stack: alStack, logs: [] });
    }
  }

  if (action !== 'getRows') {
    return jsonOut({
      ok: true,
      hint:
        'Supported actions: getRows, getPeriodData, getPeriodDataList, getActionLog, debugLastError, getLog'
    });
  }

  try {
    var sheet = getDataSheet();
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
  } catch (err) {
    var msg = String(err.message || err);
    var stack = String(err.stack || '');
    Logger.log(stack || msg);
    writeDebugLog('doGet getRows failed', {
      action: 'doGet',
      error: msg,
      stack: stack
    });
    return jsonOut({ ok: false, error: msg, stack: stack });
  }
}

function sheetRowArrayFromPayload(row) {
  var arr = [];
  for (var k = 0; k < SHEET_COLUMNS.length; k++) {
    var name = SHEET_COLUMNS[k];
    var v = row[name] != null && row[name] !== '' ? row[name] : '';
    if (name === 'period') v = periodTextForSheetWrite(v);
    arr.push(v);
  }
  return arr;
}

function sortDataSheetByPeriodAndOrder(sheet) {
  ensureHeaders(sheet);
  var lastRow = sheet.getLastRow();
  var n = SHEET_COLUMNS.length;
  if (lastRow < 2) return;
  var numRows = lastRow - 1;
  sheet.getRange(2, 1, numRows, n).sort([
    { column: 1, ascending: true },
    { column: 2, ascending: true }
  ]);
}

function buildValuesMatrixFromValidRows(validRows) {
  var n = SHEET_COLUMNS.length;
  var values = [];
  for (var i = 0; i < validRows.length; i++) {
    var row = validRows[i] || {};
    var line = [];
    for (var c = 0; c < n; c++) {
      var colName = SHEET_COLUMNS[c];
      var v = row[colName];
      if (colName === 'period') {
        line.push(periodTextForSheetWrite(v));
      } else {
        line.push(v != null && v !== '' ? v : '');
      }
    }
    if (line.length !== n) {
      throw new Error('values row width ' + line.length + ' !== ' + n);
    }
    values.push(line);
  }
  return values;
}

function deleteAllRowsForPeriod(sheet, periodParam) {
  ensureHeaders(sheet);
  var period = normalizePeriodCell(periodParam);
  if (!period) return 0;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;
  var idx = headerIndex(sheet);
  var pCol = idx['period'];
  if (pCol === undefined) return 0;
  var toDelete = [];
  for (var r = 1; r < data.length; r++) {
    if (normalizePeriodCell(data[r][pCol]) === period) {
      toDelete.push(r + 1);
    }
  }
  if (!toDelete.length) return 0;
  toDelete.sort(function(a, b) {
    return b - a;
  });
  for (var i = 0; i < toDelete.length; i++) {
    sheet.deleteRow(toDelete[i]);
  }
  return toDelete.length;
}

function prepareValidReplacePeriodRows(periodParam, rows) {
  var period = normalizePeriodCell(periodParam);
  if (!period) {
    return { ok: false, error: 'period_required', deleted: 0, inserted: 0 };
  }
  if (rows == null || !Array.isArray(rows) || rows.length === 0) {
    return {
      ok: false,
      error: 'no_valid_rows',
      deleted: 0,
      inserted: 0,
      period: period
    };
  }
  var validRows = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || {};
    var order = String(row.order == null ? '' : row.order).trim();
    if (!order) continue;
    var rowPeriod = normalizePeriodCell(row.period);
    if (rowPeriod && rowPeriod !== period) {
      return {
        ok: false,
        error: 'period_mismatch',
        deleted: 0,
        inserted: 0,
        period: period
      };
    }
    row.period = period;
    validRows.push(row);
  }
  if (validRows.length === 0) {
    return {
      ok: false,
      error: 'no_valid_rows',
      deleted: 0,
      inserted: 0,
      period: period
    };
  }
  return { ok: true, period: period, validRows: validRows };
}

function replacePeriodRows(sheet, periodParam, rows) {
  ensureHeaders(sheet);

  var rowsCount = rows && Array.isArray(rows) ? rows.length : 0;

  writeDebugLog('replacePeriod: validate start', {
    action: 'replacePeriod',
    period: periodParam,
    rowsCount: rowsCount,
    data: { stage: 'validate_start' }
  });

  var prep = prepareValidReplacePeriodRows(periodParam, rows);
  if (!prep.ok) {
    writeDebugLog('replacePeriod: validation failed', {
      action: 'replacePeriod',
      period: periodParam,
      rowsCount: rowsCount,
      validRowsCount: 0,
      error: prep.error,
      data: { stage: 'validation_failed' }
    });
    return jsonOut({
      ok: false,
      error: prep.error,
      deleted: prep.deleted || 0,
      inserted: prep.inserted || 0
    });
  }

  var period = prep.period;
  var validRows = prep.validRows;
  var n = SHEET_COLUMNS.length;
  var values;
  var valuesRows = 0;
  var valuesCols = 0;

  try {
    values = buildValuesMatrixFromValidRows(validRows);
    valuesRows = values.length;
    valuesCols = valuesRows > 0 ? values[0].length : 0;
    for (var vi = 0; vi < values.length; vi++) {
      if (values[vi].length !== n) {
        throw new Error('values[' + vi + '] width ' + values[vi].length + ' !== ' + n);
      }
    }
  } catch (prepValuesErr) {
    var prepValuesMsg = String(prepValuesErr.message || prepValuesErr);
    var prepValuesStack = String(prepValuesErr.stack || '');
    writeDebugLog('replacePeriod: values prep failed', {
      action: 'replacePeriod',
      period: period,
      rowsCount: rowsCount,
      validRowsCount: validRows.length,
      error: prepValuesMsg,
      stack: prepValuesStack,
      data: { stage: 'values_prep_failed', columnsCount: n }
    });
    PropertiesService.getScriptProperties().setProperty(
      'DEBUG_LAST_ERROR',
      prepValuesMsg + '\n' + prepValuesStack
    );
    return jsonOut({
      ok: false,
      error: prepValuesMsg,
      stack: prepValuesStack,
      deleted: 0,
      inserted: 0
    });
  }

  writeDebugLog('replacePeriod: before delete', {
    action: 'replacePeriod',
    period: period,
    rowsCount: rowsCount,
    validRowsCount: validRows.length,
    data: {
      stage: 'before_delete',
      valuesRows: valuesRows,
      valuesCols: valuesCols,
      columnsCount: n
    }
  });

  var deleted = deleteAllRowsForPeriod(sheet, period);

  writeDebugLog('replacePeriod: deleted', {
    action: 'replacePeriod',
    period: period,
    rowsCount: rowsCount,
    validRowsCount: validRows.length,
    data: { stage: 'after_delete', deleted: deleted }
  });

  var startRow = 0;
  var inserted = 0;

  if (valuesRows > 0) {
    writeDebugLog('replacePeriod: insert start', {
      action: 'replacePeriod',
      period: period,
      rowsCount: rowsCount,
      validRowsCount: validRows.length,
      data: {
        stage: 'insert_start',
        validRowsCount: validRows.length,
        columnsCount: n,
        valuesRows: valuesRows,
        valuesCols: valuesCols,
        sheetLastColumn: sheet.getLastColumn(),
        sheetName: sheet.getName()
      }
    });

    try {
      startRow = sheet.getLastRow() + 1;
      ensurePeriodColumnPlainText(sheet, 1);
      sheet.getRange(startRow, 1, valuesRows, n).setValues(values);
      sheet.getRange(startRow, 1, valuesRows, 1).setNumberFormat('@');
      inserted = valuesRows;
    } catch (insertErr) {
      var insertMsg = String(insertErr.message || insertErr);
      var insertStack = String(insertErr.stack || '');
      writeDebugLog('replacePeriod: insert failed', {
        action: 'replacePeriod',
        period: period,
        rowsCount: rowsCount,
        validRowsCount: validRows.length,
        error: insertMsg,
        stack: insertStack,
        data: {
          stage: 'insert_failed',
          message: 'insert_failed',
          valuesRows: valuesRows,
          valuesCols: valuesCols,
          columnsCount: n,
          startRow: startRow
        }
      });
      PropertiesService.getScriptProperties().setProperty(
        'DEBUG_LAST_ERROR',
        insertMsg + '\n' + insertStack
      );
      return jsonOut({
        ok: false,
        error: insertMsg,
        stack: insertStack,
        deleted: deleted,
        inserted: 0
      });
    }

    writeDebugLog('replacePeriod: inserted', {
      action: 'replacePeriod',
      period: period,
      rowsCount: rowsCount,
      validRowsCount: validRows.length,
      data: {
        stage: 'inserted',
        deleted: deleted,
        inserted: inserted,
        startRow: startRow
      }
    });
  }

  try {
    writeDebugLog('replacePeriod: sort start', {
      action: 'replacePeriod',
      period: period,
      rowsCount: rowsCount,
      validRowsCount: validRows.length,
      data: { stage: 'sort_start', deleted: deleted, inserted: inserted }
    });
    sortDataSheetByPeriodAndOrder(sheet);
  } catch (sortErr) {
    var sortMsg = String(sortErr.message || sortErr);
    var sortStack = String(sortErr.stack || '');
    writeDebugLog('replacePeriod: sort failed', {
      action: 'replacePeriod',
      period: period,
      rowsCount: rowsCount,
      validRowsCount: validRows.length,
      error: sortMsg,
      stack: sortStack,
      data: { stage: 'sort_failed', deleted: deleted, inserted: inserted }
    });
    PropertiesService.getScriptProperties().setProperty(
      'DEBUG_LAST_ERROR',
      sortMsg + '\n' + sortStack
    );
    return jsonOut({
      ok: false,
      error: sortMsg,
      stack: sortStack,
      deleted: deleted,
      inserted: inserted
    });
  }

  writeDebugLog('replacePeriod: after sort', {
    action: 'replacePeriod',
    period: period,
    rowsCount: rowsCount,
    validRowsCount: validRows.length,
    data: { stage: 'after_sort', deleted: deleted, inserted: inserted }
  });

  writeDebugLog('replacePeriod: success', {
    action: 'replacePeriod',
    period: period,
    rowsCount: rowsCount,
    validRowsCount: validRows.length,
    data: { stage: 'success', deleted: deleted, inserted: inserted }
  });

  return jsonOut({
    ok: true,
    period: period,
    deleted: deleted,
    inserted: inserted,
    written: inserted,
    sorted: true
  });
}

function upsertRows(sheet, rows) {
  ensureHeaders(sheet);
  ensurePeriodColumnPlainText(sheet, 1);
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
    var period = normalizePeriodText(row.period);
    var order = String(row.order == null ? '' : row.order).trim();
    if (!order) continue;
    row.period = period;
    var rowArr = sheetRowArrayFromPayload(row);
    var found = -1;
    for (var r = 1; r < data.length; r++) {
      var pv = normalizePeriodText(data[r][pCol]);
      var ov = String(data[r][oCol] == null ? '' : data[r][oCol]).trim();
      if (pv === period && ov === order) {
        found = r + 1;
        break;
      }
    }
    if (found > 0) {
      sheet.getRange(found, 1, found, n).setValues([rowArr]);
      sheet.getRange(found, 1).setNumberFormat('@').setValue(period);
      for (var c = 0; c < n; c++) {
        data[found - 1][c] = rowArr[c];
      }
      updated++;
    } else {
      sheet.appendRow(rowArr);
      var appended = sheet.getLastRow();
      if (appended >= 1) {
        sheet.getRange(appended, 1).setNumberFormat('@').setValue(period);
      }
      var newRow = [];
      for (var z = 0; z < n; z++) newRow.push(rowArr[z]);
      data.push(newRow);
    }
  }
  return jsonOut({ ok: true, rows: rows.length, updated: updated });
}

function doPost(e) {
  e = e || {};
  var rawBodyLength =
    e.postData && e.postData.contents ? String(e.postData.contents).length : 0;

  try {
    var cfg = getSpreadsheetConfig();
    writeDebugLog('doPost start', {
      action: 'doPost',
      data: {
        rawBodyLength: rawBodyLength,
        spreadsheetId: cfg.id || '(activeSpreadsheet fallback)',
        sheetName: cfg.sheetName || '(first sheet)'
      }
    });

    var sheet = getDataSheet();
    var body;

    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      var parseMsg = String(parseErr.message || parseErr);
      writeDebugLog('doPost invalid json', {
        action: 'doPost',
        error: parseMsg,
        stack: String(parseErr.stack || ''),
        data: { rawBodyLength: rawBodyLength }
      });
      return jsonOut({ ok: false, error: 'invalid json', stack: String(parseErr.stack || '') });
    }

    var action = body && body.action ? String(body.action) : '';

    writeDebugLog('doPost parsed', {
      action: action || 'doPost',
      period: body && body.period != null ? body.period : '',
      rowsCount: body && Array.isArray(body.rows) ? body.rows.length : 0,
      data: { rawBodyLength: rawBodyLength, payloadAction: action }
    });

    if (body && body.action === 'deletePeriod') {
      return deletePeriodPost(sheet, body);
    }

    if (body && body.action === 'replacePeriod') {
      writeDebugLog('replacePeriod: doPost received', {
        action: 'replacePeriod',
        period: body.period,
        rowsCount: body.rows && Array.isArray(body.rows) ? body.rows.length : 0,
        data: { rawBodyLength: rawBodyLength, stage: 'doPost_received' }
      });
      return replacePeriodRows(sheet, body.period, body.rows);
    }

    if (body && body.action === 'savePeriodData') {
      var periodData = body.periodData;
      var parsedPd = parseSavePeriodDataPayload(periodData);
      writeDebugLog('savePeriodData: doPost received', {
        action: 'savePeriodData',
        period: body.period,
        data: {
          rawBodyLength: rawBodyLength,
          periodDataType: periodData != null ? typeof periodData : 'null',
          isArray: Array.isArray(periodData),
          hasRows: !!(periodData && periodData.rows),
          updated_by: parsedPd.ok ? parsedPd.updatedBy : '',
          format_version: parsedPd.ok ? parsedPd.formatVersion : '',
          jsonLength: parsedPd.ok ? String(parsedPd.jsonStr).length : 0
        }
      });
      if (!parsedPd.ok) {
        return jsonOut({ ok: false, error: parsedPd.error || 'invalid_periodData' });
      }
      return upsertPeriodData(
        body.period,
        parsedPd.jsonStr,
        parsedPd.updatedBy,
        parsedPd.formatVersion
      );
    }

    if (body && body.action === 'appendActionLog') {
      return appendActionLogEntry(body.logEntry || body);
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

    return jsonOut({ ok: false, error: 'unknown_action_or_payload' });
  } catch (error) {
    var msg = String(error.message || error);
    var stack = String(error.stack || '');
    Logger.log(stack || msg);
    writeDebugLog('doPost failed', {
      action: 'doPost',
      error: msg,
      stack: stack,
      data: { rawBodyLength: rawBodyLength }
    });
    return jsonOut({ ok: false, error: msg, stack: stack });
  }
}

/**
 * Ручной тест записи в лист. Запускать из редактора Apps Script.
 * Использует period TEST-9999 — не трогает рабочие периоды.
 */
function testReplacePeriodSmall() {
  var sheet = getDataSheet();
  var testPeriod = 'TEST-9999';
  var rows = [
    {
      period: testPeriod,
      order: 'TEST ORDER',
      shipment: 0,
      materials: 1,
      processing: 0,
      other: 0,
      salary: 0,
      account25: 0,
      account26: 0,
      account44: 0,
      total_cost: 1,
      writeoff: 0,
      profit: 0,
      profit_percent: 0,
      status: 'draft'
    }
  ];

  writeDebugLog('testReplacePeriodSmall start', {
    action: 'testReplacePeriodSmall',
    period: testPeriod,
    rowsCount: 1,
    validRowsCount: 1
  });

  var result;
  var text;
  try {
    result = replacePeriodRows(sheet, testPeriod, rows);
    text = result.getContent();
  } catch (testErr) {
    var testMsg = String(testErr.message || testErr);
    var testStack = String(testErr.stack || '');
    writeDebugLog('testReplacePeriodSmall failed', {
      action: 'testReplacePeriodSmall',
      period: testPeriod,
      rowsCount: 1,
      validRowsCount: 1,
      error: testMsg,
      stack: testStack
    });
    PropertiesService.getScriptProperties().setProperty(
      'DEBUG_LAST_ERROR',
      testMsg + '\n' + testStack
    );
    Logger.log('testReplacePeriodSmall failed: ' + testMsg);
    throw testErr;
  }

  writeDebugLog('testReplacePeriodSmall done', {
    action: 'testReplacePeriodSmall',
    period: testPeriod,
    rowsCount: 1,
    validRowsCount: 1,
    data: { result: text }
  });

  Logger.log(text);
}
