// ========================================
// 💾 出張申請データ読み書き
// ========================================

var TRIP_HEADERS = [
  '出張申請ID', '申請日', '申請者Email', '申請者名', '出張開始日', '直行', '出張終了日', '直帰',
  '出張先', '目的', '交通手段', '宿泊先', '宿泊代金', '仮払金', '備考',
  'ステータス', '精算状況', '精算ID', '承認者Email', '承認日時', '差戻し理由', '更新日時',
  '経路ID', '現在ステップ', '総ステップ数', '現在ステップ名',
  '共通カレンダーイベントID', '事業所カレンダーイベントID', 'カレンダー事業所'
];

var TRIP_HEADER_ALIASES = {
  '宿泊先': ['宿泊予定'],
  '仮払金': ['概算費用']
};

var HISTORY_HEADERS = ['出張申請ID', '操作日時', '操作者Email', '操作', 'コメント'];

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getTripColumnMap_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), TRIP_HEADERS.length);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headers = headerRow.map(function(h) { return String(h || '').trim(); });
  var map = {};
  TRIP_HEADERS.forEach(function(name) {
    var idx = headers.indexOf(name);
    if (idx >= 0) map[name] = idx;
  });
  Object.keys(TRIP_HEADER_ALIASES).forEach(function(name) {
    if (map.hasOwnProperty(name)) return;
    var aliases = TRIP_HEADER_ALIASES[name];
    for (var i = 0; i < aliases.length; i++) {
      var aliasIdx = headers.indexOf(aliases[i]);
      if (aliasIdx >= 0) { map[name] = aliasIdx; break; }
    }
  });
  return map;
}

function tripCell_(data, colMap, name, defaultValue) {
  if (!colMap.hasOwnProperty(name)) return defaultValue;
  return data[colMap[name]];
}

function readTripRows_(filterFn) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_TRIPS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var colMap = getTripColumnMap_(sheet);
  var lastCol = Math.max(sheet.getLastColumn(), TRIP_HEADERS.length);
  var data = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var id = String(tripCell_(data[i], colMap, '出張申請ID', '')).trim();
    if (!id) continue;
    var row = mapTripRow_(data[i], colMap);
    if (!filterFn || filterFn(row)) rows.push(row);
  }
  return rows;
}

function mapTripRow_(data, colMap) {
  colMap = colMap || {};
  var directDepart = normalizeFlag(tripCell_(data, colMap, '直行', false));
  var directReturn = normalizeFlag(tripCell_(data, colMap, '直帰', false));
  var tripStartRaw = tripCell_(data, colMap, '出張開始日', '');
  var tripEndRaw = tripCell_(data, colMap, '出張終了日', '');
  return {
    tripRequestId: String(tripCell_(data, colMap, '出張申請ID', '')).trim(),
    requestDate: normalizeDate(tripCell_(data, colMap, '申請日', '')),
    applicantEmail: String(tripCell_(data, colMap, '申請者Email', '')).trim().toLowerCase(),
    applicantName: String(tripCell_(data, colMap, '申請者名', '')),
    tripStart: directDepart ? normalizeDate(tripStartRaw) : normalizeDateTime(tripStartRaw),
    directDepart: directDepart,
    tripEnd: directReturn ? normalizeDate(tripEndRaw) : normalizeDateTime(tripEndRaw),
    directReturn: directReturn,
    destination: String(tripCell_(data, colMap, '出張先', '')),
    purpose: String(tripCell_(data, colMap, '目的', '')),
    transport: String(tripCell_(data, colMap, '交通手段', '')),
    lodgingDestination: String(tripCell_(data, colMap, '宿泊先', '')),
    lodgingCost: normalizeAmount(tripCell_(data, colMap, '宿泊代金', 0)),
    advancePayment: normalizeAmount(tripCell_(data, colMap, '仮払金', 0)),
    note: String(tripCell_(data, colMap, '備考', '')),
    status: String(tripCell_(data, colMap, 'ステータス', TRIP_STATUS.DRAFT)).trim() || TRIP_STATUS.DRAFT,
    settlementStatus: String(tripCell_(data, colMap, '精算状況', SETTLEMENT_STATUS.NONE)).trim() || SETTLEMENT_STATUS.NONE,
    expenseClaimId: String(tripCell_(data, colMap, '精算ID', '')),
    approverEmail: String(tripCell_(data, colMap, '承認者Email', '')).trim().toLowerCase(),
    approvedAt: formatDateTime(tripCell_(data, colMap, '承認日時', '')),
    rejectReason: String(tripCell_(data, colMap, '差戻し理由', '')),
    updatedAt: formatDateTime(tripCell_(data, colMap, '更新日時', '')),
    routeId: String(tripCell_(data, colMap, '経路ID', '')).trim(),
    currentStep: parseInt(tripCell_(data, colMap, '現在ステップ', 0), 10) || 0,
    totalSteps: parseInt(tripCell_(data, colMap, '総ステップ数', 0), 10) || 0,
    currentStepName: String(tripCell_(data, colMap, '現在ステップ名', '')).trim(),
    sharedCalendarEventId: String(tripCell_(data, colMap, '共通カレンダーイベントID', '')).trim(),
    officeCalendarEventId: String(tripCell_(data, colMap, '事業所カレンダーイベントID', '')).trim(),
    calendarOffice: String(tripCell_(data, colMap, 'カレンダー事業所', '')).trim()
  };
}

function tripRowToValues_(r) {
  return [
    r.tripRequestId, r.requestDate, r.applicantEmail, r.applicantName,
    r.tripStart, r.directDepart, r.tripEnd, r.directReturn,
    r.destination, r.purpose, r.transport, r.lodgingDestination,
    r.lodgingCost, r.advancePayment, r.note, r.status, r.settlementStatus, r.expenseClaimId,
    r.approverEmail, r.approvedAt, r.rejectReason, r.updatedAt,
    r.routeId || '', r.currentStep || 0, r.totalSteps || 0, r.currentStepName || '',
    r.sharedCalendarEventId || '', r.officeCalendarEventId || '', r.calendarOffice || ''
  ];
}

function findSheetRowByFirstColumnId_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var target = String(id || '').trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === target) return i + 2;
  }
  return -1;
}

function writeTripRow_(trip) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_TRIPS) || getSpreadsheet_().insertSheet(SHEET_TRIPS);
  ensureHeaders_(sheet, TRIP_HEADERS);
  var values = tripRowToValues_(trip);
  var rowIndex = findSheetRowByFirstColumnId_(sheet, trip.tripRequestId);
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, TRIP_HEADERS.length).setValues([values]);
    return;
  }
  var nextRow = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(nextRow, 1, 1, TRIP_HEADERS.length).setValues([values]);
}

function writeAllTripRows_(sheet, rows) {
  sheet.getRange(1, 1, 1, TRIP_HEADERS.length).setValues([TRIP_HEADERS]);
  if (!rows || rows.length === 0) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    return;
  }
  var values = rows.map(tripRowToValues_);
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  sheet.getRange(2, 1, values.length, TRIP_HEADERS.length).setValues(values);
}

function buildTripRequest_(tripRequestId) {
  var rows = readTripRows_(function(r) { return r.tripRequestId === tripRequestId; });
  return rows.length ? rows[0] : null;
}

function appendHistory_(tripRequestId, action, comment) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_HISTORY) || getSpreadsheet_().insertSheet(SHEET_HISTORY);
  ensureHeaders_(sheet, HISTORY_HEADERS);
  sheet.appendRow([tripRequestId, formatDateTime(new Date()), getCurrentUserEmail_(), action, comment || '']);
}

function ensureHeaders_(sheet, headers) {
  var existing = sheet.getLastRow() >= 1 ? sheet.getRange(1, 1, 1, headers.length).getValues()[0] : [];
  var match = true;
  for (var i = 0; i < headers.length; i++) {
    if (String(existing[i] || '').trim() !== headers[i]) { match = false; break; }
  }
  if (!match) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e2e8f0');
    sheet.setFrozenRows(1);
  }
}
