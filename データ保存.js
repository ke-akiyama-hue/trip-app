// ========================================
// 💾 出張申請データ読み書き
// ========================================

var TRIP_HEADERS = [
  '出張申請ID', '申請日', '申請者Email', '申請者名', '出張開始日', '出張終了日',
  '出張先', '目的', '交通手段', '宿泊予定', '概算費用', '備考',
  'ステータス', '精算状況', '精算ID', '承認者Email', '承認日時', '差戻し理由', '更新日時',
  '経路ID', '現在ステップ', '総ステップ数', '現在ステップ名'
];

var HISTORY_HEADERS = ['出張申請ID', '操作日時', '操作者Email', '操作', 'コメント'];

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function readTripRows_(filterFn) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_TRIPS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow(), TRIP_HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = mapTripRow_(data[i]);
    if (!filterFn || filterFn(row)) rows.push(row);
  }
  return rows;
}

function mapTripRow_(data) {
  return {
    tripRequestId: String(data[0]),
    requestDate: normalizeDate(data[1]),
    applicantEmail: String(data[2] || '').trim().toLowerCase(),
    applicantName: String(data[3] || ''),
    tripStart: normalizeDate(data[4]),
    tripEnd: normalizeDate(data[5]),
    destination: String(data[6] || ''),
    purpose: String(data[7] || ''),
    transport: String(data[8] || ''),
    lodgingPlan: String(data[9] || ''),
    estimatedCost: normalizeAmount(data[10]),
    note: String(data[11] || ''),
    status: String(data[12] || TRIP_STATUS.DRAFT),
    settlementStatus: String(data[13] || SETTLEMENT_STATUS.NONE),
    expenseClaimId: String(data[14] || ''),
    approverEmail: String(data[15] || '').trim().toLowerCase(),
    approvedAt: formatDateTime(data[16]),
    rejectReason: String(data[17] || ''),
    updatedAt: formatDateTime(data[18]),
    routeId: String(data[19] || '').trim(),
    currentStep: parseInt(data[20], 10) || 0,
    totalSteps: parseInt(data[21], 10) || 0,
    currentStepName: String(data[22] || '').trim()
  };
}

function tripRowToValues_(r) {
  return [
    r.tripRequestId, r.requestDate, r.applicantEmail, r.applicantName,
    r.tripStart, r.tripEnd, r.destination, r.purpose, r.transport, r.lodgingPlan,
    r.estimatedCost, r.note, r.status, r.settlementStatus, r.expenseClaimId,
    r.approverEmail, r.approvedAt, r.rejectReason, r.updatedAt,
    r.routeId || '', r.currentStep || 0, r.totalSteps || 0, r.currentStepName || ''
  ];
}

function writeTripRow_(trip) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_TRIPS) || getSpreadsheet_().insertSheet(SHEET_TRIPS);
  ensureHeaders_(sheet, TRIP_HEADERS);
  var all = readTripRows_();
  var idx = -1;
  for (var i = 0; i < all.length; i++) {
    if (all[i].tripRequestId === trip.tripRequestId) { idx = i; break; }
  }
  if (idx >= 0) all[idx] = trip;
  else all.push(trip);
  writeAllTripRows_(sheet, all);
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
