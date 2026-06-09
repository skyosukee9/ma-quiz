var REPORT_SHEET_NAME = "完答レポート";
var PERSONAL_TEMPLATE_SHEET_NAME = "接続テスト_人物別";
var MIN_FORMAT_ROWS = 100;
var REPORT_HEADERS = ["名前", "回答日時", "所要時間", "大問名", "正答率", "全体進捗率", "結果サマリー", "分析レポート"];

function doPost(e) {
  var data = JSON.parse((e.postData && e.postData.contents) || "{}");
  recordReport_(data);
  return output_({ ok: true });
}

function doGet(e) {
  var data = (e && e.parameter) || {};
  if (data.action === "names") return output_({ names: getKnownNames_() }, data.callback);
  if (data.playerName || data.genreTitle) recordReport_(data);
  return output_({ ok: true }, data.callback);
}

function recordReport_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, REPORT_SHEET_NAME);
  ensureHeaders_(sheet);
  var row = [
    data.playerName || "",
    new Date(),
    data.elapsedLabel || "",
    data.genreTitle || "",
    formatPercent_(data.percent),
    calculateProgress_(sheet, data),
    data.reportSummary || "",
    data.reportAnalysis || ""
  ];
  sheet.appendRow(row);
  applyLayout_(ss, sheet);
  appendPersonalReport_(ss, data.playerName, row);
}

function calculateProgress_(sheet, data) {
  var name = String(data.playerName || "").trim();
  var genre = String(data.genreTitle || "").trim();
  var total = Number(data.totalTopics || 0);
  if (!name || !genre || !isFinite(total) || total <= 0) return "";
  var done = {};
  done[genre] = true;
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, REPORT_HEADERS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0] || "").trim() === name && values[i][3]) {
        done[String(values[i][3]).trim()] = true;
      }
    }
  }
  return Math.min(100, Math.round((Object.keys(done).length / total) * 100)) + "%";
}

function getKnownNames_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, REPORT_SHEET_NAME);
  ensureHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var seen = {};
  var names = [];
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0] || "").trim();
    if (name && !seen[name]) {
      seen[name] = true;
      names.push(name);
    }
  }
  return names.sort(function(a, b) { return a.localeCompare(b, "ja"); });
}

function appendPersonalReport_(ss, playerName, row) {
  var sheetName = getPersonalSheetName_(playerName);
  if (!sheetName) return;
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  ensureHeaders_(sheet);
  sheet.appendRow(row);
  applyLayout_(ss, sheet);
}

function applyLayout_(ss, sheet) {
  var template = ss.getSheetByName(PERSONAL_TEMPLATE_SHEET_NAME);
  var cols = Math.min(REPORT_HEADERS.length, sheet.getMaxColumns());
  var rows = Math.min(Math.max(sheet.getLastRow(), MIN_FORMAT_ROWS), sheet.getMaxRows());
  if (template && template.getSheetId() !== sheet.getSheetId()) {
    var tCols = Math.min(cols, template.getMaxColumns());
    var tRows = Math.min(rows, template.getMaxRows());
    for (var c = 1; c <= tCols; c++) sheet.setColumnWidth(c, template.getColumnWidth(c));
    for (var r = 1; r <= tRows; r++) sheet.setRowHeight(r, template.getRowHeight(r));
    template.getRange(1, 1, tRows, tCols).copyTo(sheet.getRange(1, 1, tRows, tCols), { formatOnly: true });
    sheet.setFrozenRows(template.getFrozenRows());
    sheet.setFrozenColumns(template.getFrozenColumns());
  }
  sheet.getRange(1, 1, rows, cols).setWrap(true).setVerticalAlignment("middle");
  sheet.autoResizeRows(1, Math.max(sheet.getLastRow(), 1));
}

function getPersonalSheetName_(playerName) {
  var name = String(playerName || "").trim();
  if (!name) return "";
  return name.replace(/[\\/?*[\]:]/g, " ").replace(/\s+/g, " ").trim().slice(0, 90) || "未設定";
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders_(sheet) {
  var maxColumns = Math.max(sheet.getMaxColumns(), REPORT_HEADERS.length);
  sheet.getRange(1, 1, 1, maxColumns).clearContent();
  sheet.getRange(1, 1, 1, REPORT_HEADERS.length).setValues([REPORT_HEADERS]);
  sheet.setFrozenRows(1);
}

function formatPercent_(value) {
  var number = Number(value || 0);
  return isFinite(number) ? Math.round(number) + "%" : "";
}

function output_(payload, callback) {
  if (callback && /^[\w.$]+$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(payload) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
