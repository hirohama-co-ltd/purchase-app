// ========================================
// ⚙️ 購買申請アプリ 共通設定
// ========================================

var MASTER_SS_ID = '1FrxPVUeKecY8SXwc5daMxjGT0MzQKZ_toa77PfO4iQo';
var EMPLOYEE_MASTER_SHEET_NAME = '社員マスタ';
var WORKFLOW_SS_ID = '19zhtLt23UOpysCpbwH9X-gom5ohVKbW2nARECDvCfIk';
var APP_CODE = 'PURCHASE_REQUEST';

/** 申請ポータルのWebアプリURL（戻る導線用） */
var PORTAL_URL = 'https://script.google.com/macros/s/AKfycbwLx0zRZApqzd9d3Np8HhMQJzOzp1L_TSvL4xiL_Svrwiguyuk1oLQAcAlSX8F3OGc8/exec';

var PURCHASE_STATUS = {
  DRAFT: '下書き',
  SUBMITTED: '申請中',
  APPROVED: '承認済',
  REJECTED: '差戻し',
  WITHDRAWN: '取り下げ',
  CANCELLED: '取消'
};

var SHEET_PURCHASES = '購買申請一覧';
var SHEET_PURCHASE_DETAILS = '購買明細';
var SHEET_SUPPLIER_MASTER = '購入先マスタ';
var SHEET_MAKER_MASTER = 'メーカーマスタ';
var SHEET_UNREGISTERED_MASTER_CANDIDATES = '未登録マスタ候補';
var SHEET_HISTORY = '承認履歴';
var MASTER_CACHE_TTL_SEC = 600;

var PURCHASE_MASTER_STATUS = {
  REGISTERED: '登録済',
  PENDING: 'マスタ未登録あり'
};

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

function generatePurchaseRequestId_() {
  var tz = Session.getScriptTimeZone();
  var prefix = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  var rand = Math.floor(Math.random() * 9000) + 1000;
  return 'PR-' + prefix + '-' + rand;
}

function getCurrentUserEmail_() {
  try {
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
