const LOG_SHEET_NAME = "回答ログ";
const SUMMARY_SHEET_NAME = "人物別サマリー";

const LOG_HEADERS = [
  "記録日時",
  "名前",
  "パート",
  "大問ID",
  "大問名",
  "出題モード",
  "出題数",
  "正解数",
  "正答率",
  "回答時間秒",
  "回答時間",
  "結果",
  "間違えた問題数",
  "間違えた問題",
  "間違えたタグ",
  "ブラウザID",
  "結果サマリー",
  "分析レポート",
  "重点復習タグ"
];

const SUMMARY_HEADERS = [
  "名前",
  "受験回数",
  "クリア回数",
  "平均正答率",
  "平均回答時間",
  "最終受験日時",
  "最終パート",
  "最終大問",
  "最新レポート"
];

function doPost(e) {
  const data = JSON.parse((e.postData && e.postData.contents) || "{}");
  recordProgress_(data);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const data = (e && e.parameter) || {};
  if (data.playerName || data.genreId) {
    recordProgress_(data);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function recordProgress_(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateSheet_(spreadsheet, LOG_SHEET_NAME);
  ensureHeaders_(logSheet, LOG_HEADERS);

  logSheet.appendRow([
    new Date(),
    data.playerName || "",
    data.part || "",
    data.genreId || "",
    data.genreTitle || "",
    data.mode || "",
    toNumber_(data.total),
    toNumber_(data.correctCount),
    toNumber_(data.percent),
    toNumber_(data.elapsedSeconds),
    data.elapsedLabel || "",
    data.cleared === true || data.cleared === "true" ? "クリア" : "未クリア",
    toNumber_(data.wrongCount),
    data.wrongQuestions || "",
    data.wrongTags || "",
    data.browserId || "",
    data.reportSummary || "",
    data.reportAnalysis || "",
    data.focusTags || ""
  ]);

  updateSummary_(spreadsheet);
}

function getOrCreateSheet_(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = headers.some((header, index) => currentHeaders[index] !== header);
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function updateSummary_(spreadsheet) {
  const logSheet = getOrCreateSheet_(spreadsheet, LOG_SHEET_NAME);
  ensureHeaders_(logSheet, LOG_HEADERS);

  const summarySheet = getOrCreateSheet_(spreadsheet, SUMMARY_SHEET_NAME);
  ensureHeaders_(summarySheet, SUMMARY_HEADERS);

  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const rows = logSheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  const people = new Map();

  rows.forEach((row) => {
    const name = String(row[1] || "").trim();
    if (!name) {
      return;
    }

    const current = people.get(name) || {
      attempts: 0,
      clears: 0,
      percentTotal: 0,
      secondsTotal: 0,
      lastAt: null,
      lastPart: "",
      lastGenre: "",
      lastReport: ""
    };

    const recordedAt = row[0] instanceof Date ? row[0] : new Date(row[0]);
    current.attempts += 1;
    current.clears += row[11] === "クリア" ? 1 : 0;
    current.percentTotal += Number(row[8] || 0);
    current.secondsTotal += Number(row[9] || 0);

    if (!current.lastAt || recordedAt > current.lastAt) {
      current.lastAt = recordedAt;
      current.lastPart = row[2] || "";
      current.lastGenre = row[4] || "";
      current.lastReport = row[17] || row[16] || "";
    }

    people.set(name, current);
  });

  const summaryRows = Array.from(people.entries())
    .sort((left, right) => left[0].localeCompare(right[0], "ja"))
    .map(([name, value]) => [
      name,
      value.attempts,
      value.clears,
      `${Math.round(value.percentTotal / value.attempts)}%`,
      formatSeconds_(Math.round(value.secondsTotal / value.attempts)),
      value.lastAt,
      value.lastPart,
      value.lastGenre,
      value.lastReport
    ]);

  if (summarySheet.getLastRow() > 1) {
    summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, SUMMARY_HEADERS.length).clearContent();
  }
  summarySheet.getRange(2, 1, summaryRows.length, SUMMARY_HEADERS.length).setValues(summaryRows);
  summarySheet.autoResizeColumns(1, SUMMARY_HEADERS.length);
}

function formatSeconds_(seconds) {
  const safeSeconds = Math.max(0, seconds);
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const rest = safeSeconds % 60;
  const paddedMinutes = String(minutes).padStart(2, "0");
  const paddedSeconds = String(rest).padStart(2, "0");

  return hours > 0
    ? `${hours}:${paddedMinutes}:${paddedSeconds}`
    : `${paddedMinutes}:${paddedSeconds}`;
}

function toNumber_(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) ? number : 0;
}
