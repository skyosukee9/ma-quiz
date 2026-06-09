const REPORT_SHEET_NAME = "完答レポート";
const PERSONAL_TEMPLATE_SHEET_NAME = "接続テスト_人物別";

const REPORT_HEADERS = [
  "名前",
  "回答日時",
  "所要時間",
  "大問名",
  "正答率",
  "全体進捗率",
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
  if (data.action === "names") {
    return createOutput_({ names: getKnownNames_() }, data.callback);
  }

  if (data.playerName || data.genreTitle) {
    recordReport_(data);
  }

  return createOutput_({ ok: true }, data.callback);
}

function recordReport_(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(spreadsheet, REPORT_SHEET_NAME);
  ensureHeaders_(sheet, REPORT_HEADERS);
  const progress = calculateProgress_(sheet, data);
  const row = [
    data.playerName || "",
    new Date(),
    data.elapsedLabel || "",
    data.genreTitle || "",
    formatPercent_(data.percent),
    progress,
    data.reportSummary || "",
    data.reportAnalysis || ""
  ];

  sheet.appendRow(row);
  sheet.autoResizeColumns(1, REPORT_HEADERS.length);
  appendPersonalReport_(spreadsheet, data.playerName, row);
}

function calculateProgress_(sheet, data) {
  const name = String(data.playerName || "").trim();
  const genreTitle = String(data.genreTitle || "").trim();
  const totalTopics = Number(data.totalTopics || 0);
  if (!name || !genreTitle || !Number.isFinite(totalTopics) || totalTopics <= 0) {
    return "";
  }

  const completedTopics = new Set([genreTitle]);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, REPORT_HEADERS.length).getValues();
    values.forEach((row) => {
      const rowName = String(row[0] || "").trim();
      const rowGenreTitle = String(row[3] || "").trim();
      if (rowName === name && rowGenreTitle) {
        completedTopics.add(rowGenreTitle);
      }
    });
  }

  return `${Math.min(100, Math.round((completedTopics.size / totalTopics) * 100))}%`;
}

function getKnownNames_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(spreadsheet, REPORT_SHEET_NAME);
  ensureHeaders_(sheet, REPORT_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const names = values
    .map((row) => String(row[0] || "").trim())
    .filter(Boolean);
  return [...new Set(names)].sort((left, right) => left.localeCompare(right, "ja"));
}

function createOutput_(payload, callback) {
  if (callback && /^[\w.$]+$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function appendPersonalReport_(spreadsheet, playerName, row) {
  const sheetName = getPersonalSheetName_(playerName);
  if (!sheetName) {
    return;
  }

  const sheet = getOrCreatePersonalSheet_(spreadsheet, sheetName);
  ensureHeaders_(sheet, REPORT_HEADERS);
  sheet.appendRow(row);
}

function getOrCreatePersonalSheet_(spreadsheet, sheetName) {
  const existing = spreadsheet.getSheetByName(sheetName);
  if (existing) {
    return existing;
  }

  const sheet = spreadsheet.insertSheet(sheetName);
  applyPersonalSheetTemplate_(spreadsheet, sheet);
  return sheet;
}

function applyPersonalSheetTemplate_(spreadsheet, sheet) {
  const template = spreadsheet.getSheetByName(PERSONAL_TEMPLATE_SHEET_NAME);
  if (!template || template.getSheetId() === sheet.getSheetId()) {
    return;
  }

  const columnCount = Math.min(REPORT_HEADERS.length, template.getMaxColumns(), sheet.getMaxColumns());
  const rowCount = Math.min(Math.max(template.getLastRow(), 20), template.getMaxRows(), sheet.getMaxRows());

  for (let column = 1; column <= columnCount; column += 1) {
    sheet.setColumnWidth(column, template.getColumnWidth(column));
  }

  for (let row = 1; row <= rowCount; row += 1) {
    sheet.setRowHeight(row, template.getRowHeight(row));
  }

  template
    .getRange(1, 1, rowCount, columnCount)
    .copyTo(sheet.getRange(1, 1, rowCount, columnCount), { formatOnly: true });

  sheet.setFrozenRows(template.getFrozenRows());
  sheet.setFrozenColumns(template.getFrozenColumns());
}

function getPersonalSheetName_(playerName) {
  const name = String(playerName || "").trim();
  if (!name) {
    return "";
  }

  const sanitized = name.replace(/[\\/?*[\]:]/g, " ").replace(/\s+/g, " ").trim();
  return sanitized.slice(0, 90) || "未設定";
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
