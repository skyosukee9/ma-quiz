const REPORT_SHEET_NAME = "完答レポート";

const REPORT_HEADERS = [
  "名前",
  "回答日時",
  "所要時間",
  "大問名",
  "正答率",
  "結果サマリー",
  "分析レポート"
];

function doPost(e) {
  const data = JSON.parse((e.postData && e.postData.contents) || "{}");
  recordReport_(data);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const data = (e && e.parameter) || {};
  if (data.playerName || data.genreTitle) {
    recordReport_(data);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function recordReport_(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(spreadsheet, REPORT_SHEET_NAME);
  ensureHeaders_(sheet, REPORT_HEADERS);

  sheet.appendRow([
    data.playerName || "",
    new Date(),
    data.elapsedLabel || "",
    data.genreTitle || "",
    formatPercent_(data.percent),
    data.reportSummary || "",
    data.reportAnalysis || ""
  ]);

  sheet.autoResizeColumns(1, REPORT_HEADERS.length);
}

function getOrCreateSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  const maxColumns = Math.max(sheet.getMaxColumns(), headers.length);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  sheet.getRange(1, 1, 1, maxColumns).clearContent();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function formatPercent_(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) ? `${Math.round(number)}%` : "";
}
