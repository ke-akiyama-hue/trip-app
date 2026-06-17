// ========================================
// ⚙️ 出張申請アプリ 共通設定
// ========================================

var MASTER_SS_ID = '1FrxPVUeKecY8SXwc5daMxjGT0MzQKZ_toa77PfO4iQo';
var EMPLOYEE_MASTER_SHEET_NAME = '社員マスタ';

/** ワークフロー設定アプリのスプレッドシートID（clasp push 後に URL から設定） */
var WORKFLOW_SS_ID = '19zhtLt23UOpysCpbwH9X-gom5ohVKbW2nARECDvCfIk';

/** このアプリのワークフロー識別コード */
var APP_CODE = 'TRIP_REQUEST';

/** 出張申請ステータス */
var TRIP_STATUS = {
  DRAFT: '下書き',
  SUBMITTED: '申請中',
  APPROVED: '承認済',
  REJECTED: '差戻し',
  CANCELLED: '取消'
};

/** 精算連携ステータス（出張旅費精算アプリが更新） */
var SETTLEMENT_STATUS = {
  NONE: '未精算',
  IN_PROGRESS: '精算中',
  DONE: '精算完了'
};

var SHEET_TRIPS = '出張申請一覧';
var SHEET_HISTORY = '承認履歴';

var MASTER_CACHE_TTL_SEC = 600;

function normalizeDate(dateInput) {
  var tz = Session.getScriptTimeZone();
  if (!dateInput) return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
    return Utilities.formatDate(dateInput, tz, 'yyyy-MM-dd');
  }
  var s = String(dateInput).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) return s.replace(/\//g, '-');
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

function formatDateTime(val) {
  if (!val) return '';
  var tz = Session.getScriptTimeZone();
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, tz, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(val);
}

function normalizeAmount(val) {
  var n = parseInt(String(val || '0').replace(/[,，]/g, ''), 10);
  return isNaN(n) ? 0 : Math.max(0, n);
}

function generateTripRequestId_() {
  var tz = Session.getScriptTimeZone();
  var prefix = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var rand = Math.floor(Math.random() * 9000) + 1000;
  return 'BT-' + prefix + '-' + rand;
}

function getCurrentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

function getCachedJson_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    if (raw) return JSON.parse(raw);
  } catch (e) { /* ignore */ }
  return null;
}

function putCachedJson_(key, value, expirationInSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), expirationInSeconds || MASTER_CACHE_TTL_SEC);
  } catch (e) { /* ignore */ }
}

function clearMasterCaches_() {
  try {
    if (MASTER_SS_ID) CacheService.getScriptCache().remove('employees_' + MASTER_SS_ID);
    if (WORKFLOW_SS_ID) {
      CacheService.getScriptCache().remove('wf_routes_' + WORKFLOW_SS_ID);
      CacheService.getScriptCache().remove('wf_steps_' + WORKFLOW_SS_ID);
      CacheService.getScriptCache().remove('wf_bindings_' + WORKFLOW_SS_ID);
    }
  } catch (e) { /* ignore */ }
}
