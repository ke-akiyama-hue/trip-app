// ========================================
// ⚙️ 出張申請アプリ 共通設定
// ========================================

var MASTER_SS_ID = '1FrxPVUeKecY8SXwc5daMxjGT0MzQKZ_toa77PfO4iQo';
var EMPLOYEE_MASTER_SHEET_NAME = '社員マスタ';

/** ワークフロー設定アプリのスプレッドシートID（clasp push 後に URL から設定） */
var WORKFLOW_SS_ID = '19zhtLt23UOpysCpbwH9X-gom5ohVKbW2nARECDvCfIk';

/** このアプリのワークフロー識別コード */
var APP_CODE = 'TRIP_REQUEST';

/** 申請ポータルのWebアプリURL（戻る導線用・デプロイ後の /exec URL を設定） */
var PORTAL_URL = 'https://script.google.com/macros/s/AKfycbwLx0zRZApqzd9d3Np8HhMQJzOzp1L_TSvL4xiL_Svrwiguyuk1oLQAcAlSX8F3OGc8/exec';

/** 出張申請ステータス */
var TRIP_STATUS = {
  DRAFT: '下書き',
  SUBMITTED: '申請中',
  APPROVED: '承認済',
  REJECTED: '差戻し',
  WITHDRAWN: '取り下げ',
  CANCELLED: '取消'
};

/** 出張日程を反映する共通GoogleカレンダーID（未設定なら共通カレンダー同期なし） */
var TRIP_SHARED_CALENDAR_ID = 'c_917c7ef0199eb4164e5f167f695fd87c3fcf2941502940abdc59eef9c216180e@group.calendar.google.com';

/**
 * 申請者の事業所ごとに反映するGoogleカレンダーID。
 * 社員マスタの「事業所」列の値とキーを一致させる。
 * 例: { '本社': 'xxxxx@group.calendar.google.com', '大阪': 'yyyyy@group.calendar.google.com' }
 */
var TRIP_OFFICE_CALENDAR_IDS = {
  '大阪工場': 'c_7307233b51a453a13eeb28ef8cdcd92c41e7e6d76ed09ae6e3a912ab22ff4aec@group.calendar.google.com',
  '大阪営業所': 'c_26c0d7b1d33ebeab52763026e0f7044b0bda669fdee00ad8642804ab4cc8b6be@group.calendar.google.com'
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

function normalizeDateTime(dateInput) {
  var tz = Session.getScriptTimeZone();
  if (!dateInput) return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
    return Utilities.formatDate(dateInput, tz, 'yyyy-MM-dd HH:mm');
  }
  var s = String(dateInput).trim();
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}/.test(s)) return s.substring(0, 16);
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(s)) return s.replace('T', ' ').substring(0, 16);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s + ' 09:00';
  if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) return s.replace(/\//g, '-') + ' 09:00';
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm');
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
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

function normalizeFlag(val) {
  return val === true || val === 'TRUE' || val === 'はい' || val === '1' || val === 1;
}

function generateTripRequestId_() {
  var tz = Session.getScriptTimeZone();
  var prefix = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var rand = Math.floor(Math.random() * 9000) + 1000;
  return 'BT-' + prefix + '-' + rand;
}

function getCurrentUserEmail_() {
  try {
    // executeAs=USER_DEPLOYING でも、申請者・承認者の判定はログインユーザーで行う。
    // EffectiveUser はデプロイ実行者になるため、ここでは使わない。
    var email = Session.getActiveUser().getEmail() || '';
    return String(email).trim().toLowerCase();
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
