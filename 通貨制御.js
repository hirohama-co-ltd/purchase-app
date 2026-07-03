// ========================================
// 💱 購買申請 — 通貨・為替レート
// ========================================

var DEFAULT_PURCHASE_CURRENCY = 'JPY';
var FX_SHEET_NAME = '_為替レート';

var PURCHASE_CURRENCIES = [
  { code: 'JPY', label: '日本円', symbol: '¥' },
  { code: 'USD', label: '米ドル', symbol: '$' },
  { code: 'EUR', label: 'ユーロ', symbol: '€' },
  { code: 'GBP', label: '英ポンド', symbol: '£' },
  { code: 'CNY', label: '中国元', symbol: 'CN¥' }
];

var FX_CACHE_TTL_SEC = 3600;

function normalizeCurrencyCode_(code) {
  var c = String(code || DEFAULT_PURCHASE_CURRENCY).trim().toUpperCase();
  for (var i = 0; i < PURCHASE_CURRENCIES.length; i++) {
    if (PURCHASE_CURRENCIES[i].code === c) return c;
  }
  return DEFAULT_PURCHASE_CURRENCY;
}

function getCurrencyMeta_(code) {
  code = normalizeCurrencyCode_(code);
  for (var i = 0; i < PURCHASE_CURRENCIES.length; i++) {
    if (PURCHASE_CURRENCIES[i].code === code) return PURCHASE_CURRENCIES[i];
  }
  return PURCHASE_CURRENCIES[0];
}

function normalizeCurrencyAmount_(val, currency) {
  if (normalizeCurrencyCode_(currency) === 'JPY') return normalizeAmount(val);
  var n = parseFloat(String(val || '0').replace(/[,，]/g, ''));
  if (isNaN(n)) return 0;
  return Math.max(0, Math.round(n * 100) / 100);
}

function getExchangeRateSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(FX_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FX_SHEET_NAME);
    try { sheet.hideSheet(); } catch (e) { /* ignore */ }
  }
  return sheet;
}

function fetchExchangeRateToJpyViaGoogleFinance_(currencyCode) {
  currencyCode = normalizeCurrencyCode_(currencyCode);
  var sheet = getExchangeRateSheet_();
  var pair = currencyCode + 'JPY';
  sheet.getRange('A1').setValue(pair);
  sheet.getRange('B1').setFormula('=GOOGLEFINANCE("CURRENCY:' + pair + '")');
  SpreadsheetApp.flush();

  var rate = null;
  for (var attempt = 0; attempt < 6; attempt++) {
    if (attempt > 0) Utilities.sleep(400);
    rate = sheet.getRange('B1').getValue();
    if (typeof rate === 'number' && rate > 0 && !isNaN(rate)) break;
    var text = String(rate || '');
    if (text.indexOf('#') === 0) {
      return { success: false, message: '為替レートを取得できませんでした（' + text + '）。' };
    }
  }

  if (typeof rate !== 'number' || rate <= 0 || isNaN(rate)) {
    return { success: false, message: '為替レートを取得できませんでした。しばらくしてから再度お試しください。' };
  }

  return {
    success: true,
    currency: currencyCode,
    rate: Math.round(rate * 10000) / 10000,
    rateDate: normalizeDate(new Date()),
    fetchedAt: formatDateTime(new Date()),
    source: 'GOOGLEFINANCE'
  };
}

function fetchExchangeRateToJpy_(currencyCode) {
  currencyCode = normalizeCurrencyCode_(currencyCode);
  if (currencyCode === 'JPY') {
    return {
      success: true,
      currency: 'JPY',
      rate: 1,
      rateDate: normalizeDate(new Date()),
      fetchedAt: formatDateTime(new Date())
    };
  }

  var cacheKey = 'fx_' + currencyCode + '_JPY';
  var cached = getCachedJson_(cacheKey);
  if (cached && cached.rate) return cached;

  var result = fetchExchangeRateToJpyViaGoogleFinance_(currencyCode);
  if (result.success) putCachedJson_(cacheKey, result, FX_CACHE_TTL_SEC);
  return result;
}

function convertToEstimatedJpy_(amount, currency, exchangeRate) {
  currency = normalizeCurrencyCode_(currency);
  var amt = normalizeCurrencyAmount_(amount, currency);
  if (currency === 'JPY') return amt;
  var rate = parseFloat(exchangeRate);
  if (!rate || rate <= 0) return 0;
  return Math.round(amt * rate);
}

function resolvePurchaseCurrencyFields_(currency, totalAmount, payload, isSubmit, existing) {
  currency = normalizeCurrencyCode_(currency || (existing && existing.currency) || (payload && payload.currency));
  totalAmount = normalizeCurrencyAmount_(totalAmount, currency);

  if (currency === 'JPY') {
    return {
      success: true,
      currency: 'JPY',
      exchangeRate: 1,
      exchangeRateDate: '',
      estimatedJpyAmount: totalAmount
    };
  }

  var rateResult;
  if (isSubmit) {
    rateResult = fetchExchangeRateToJpy_(currency);
  } else {
    var clientRate = parseFloat(payload && payload.exchangeRate);
    if (clientRate > 0) {
      rateResult = {
        success: true,
        currency: currency,
        rate: clientRate,
        rateDate: String((payload && payload.exchangeRateDate) || '').trim()
      };
    } else {
      rateResult = fetchExchangeRateToJpy_(currency);
    }
  }
  if (!rateResult.success) return rateResult;

  return {
    success: true,
    currency: currency,
    exchangeRate: rateResult.rate,
    exchangeRateDate: rateResult.rateDate || '',
    estimatedJpyAmount: convertToEstimatedJpy_(totalAmount, currency, rateResult.rate)
  };
}

function getPurchaseCurrencyOptions_() {
  return PURCHASE_CURRENCIES.slice();
}

function getExchangeRateApi(currency) {
  try {
    return fetchExchangeRateToJpy_(currency);
  } catch (e) {
    Logger.log('getExchangeRateApi: ' + e.message);
    return { success: false, message: '為替レートの取得に失敗しました: ' + e.message };
  }
}

/** Apps Script エディタ / メニューから為替取得を確認する用 */
function authorizeExchangeRateAccess() {
  return getExchangeRateApi('USD');
}

function formatPurchaseAmountText_(amount, currency) {
  currency = normalizeCurrencyCode_(currency);
  var meta = getCurrencyMeta_(currency);
  var val = normalizeCurrencyAmount_(amount, currency);
  if (currency === 'JPY') return Number(val).toLocaleString('ja-JP') + ' 円';
  return meta.symbol + Number(val).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) +
    '（' + meta.label + '）';
}
