// ========================================
// 💾 購買申請データ読み書き
// ========================================

var PURCHASE_HEADERS = [
  '購買申請ID', '申請日', '申請者Email', '申請者名', '希望納期',
  '購入先', '品名', '数量', '単価', '合計金額', '購買目的', '予算区分', '支払方法', '備考',
  'マスタ登録状態', '未登録マスタ件数', 'ステータス', '承認者Email', '承認日時', '差戻し理由', '更新日時',
  '経路ID', '現在ステップ', '総ステップ数', '現在ステップ名'
];

var HISTORY_HEADERS = ['購買申請ID', '操作日時', '操作者Email', '操作', 'コメント'];
var PURCHASE_DETAIL_HEADERS = ['購買申請ID', '行No', 'メーカー', '品番', '品名', '数量', '単価', '金額', '備考'];
var UNREGISTERED_MASTER_HEADERS = ['候補ID', '種別', '入力名', '正式表示名', 'コード', '別名', '状態', '関連申請ID', '関連行No', '更新日時', '登録日時', '処理者Email'];

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getPurchaseColumnMap_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), PURCHASE_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  var map = {};
  PURCHASE_HEADERS.forEach(function(name) {
    var idx = headers.indexOf(name);
    if (idx >= 0) map[name] = idx;
  });
  return map;
}

function purchaseCell_(data, colMap, name, defaultValue) {
  if (!colMap.hasOwnProperty(name)) return defaultValue;
  return data[colMap[name]];
}

function readPurchaseRows_(filterFn) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var colMap = getPurchaseColumnMap_(sheet);
  var lastCol = Math.max(sheet.getLastColumn(), PURCHASE_HEADERS.length);
  var data = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var id = String(purchaseCell_(data[i], colMap, '購買申請ID', '')).trim();
    if (!id) continue;
    var row = mapPurchaseRow_(data[i], colMap);
    if (!filterFn || filterFn(row)) rows.push(row);
  }
  return rows;
}

function mapPurchaseRow_(data, colMap) {
  colMap = colMap || {};
  return {
    purchaseRequestId: String(purchaseCell_(data, colMap, '購買申請ID', '')).trim(),
    requestDate: normalizeDate(purchaseCell_(data, colMap, '申請日', '')),
    applicantEmail: String(purchaseCell_(data, colMap, '申請者Email', '')).trim().toLowerCase(),
    applicantName: String(purchaseCell_(data, colMap, '申請者名', '')),
    desiredDate: normalizeDate(purchaseCell_(data, colMap, '希望納期', '')),
    supplier: String(purchaseCell_(data, colMap, '購入先', '')),
    itemName: String(purchaseCell_(data, colMap, '品名', '')),
    quantity: normalizeAmount(purchaseCell_(data, colMap, '数量', 0)),
    unitPrice: normalizeAmount(purchaseCell_(data, colMap, '単価', 0)),
    totalAmount: normalizeAmount(purchaseCell_(data, colMap, '合計金額', 0)),
    purpose: String(purchaseCell_(data, colMap, '購買目的', '')),
    budgetCategory: String(purchaseCell_(data, colMap, '予算区分', '')),
    paymentMethod: String(purchaseCell_(data, colMap, '支払方法', '')),
    note: String(purchaseCell_(data, colMap, '備考', '')),
    masterStatus: String(purchaseCell_(data, colMap, 'マスタ登録状態', PURCHASE_MASTER_STATUS.REGISTERED)).trim() || PURCHASE_MASTER_STATUS.REGISTERED,
    unregisteredMasterCount: parseInt(purchaseCell_(data, colMap, '未登録マスタ件数', 0), 10) || 0,
    status: String(purchaseCell_(data, colMap, 'ステータス', PURCHASE_STATUS.DRAFT)).trim() || PURCHASE_STATUS.DRAFT,
    approverEmail: String(purchaseCell_(data, colMap, '承認者Email', '')).trim().toLowerCase(),
    approvedAt: formatDateTime(purchaseCell_(data, colMap, '承認日時', '')),
    rejectReason: String(purchaseCell_(data, colMap, '差戻し理由', '')),
    updatedAt: formatDateTime(purchaseCell_(data, colMap, '更新日時', '')),
    routeId: String(purchaseCell_(data, colMap, '経路ID', '')).trim(),
    currentStep: parseInt(purchaseCell_(data, colMap, '現在ステップ', 0), 10) || 0,
    totalSteps: parseInt(purchaseCell_(data, colMap, '総ステップ数', 0), 10) || 0,
    currentStepName: String(purchaseCell_(data, colMap, '現在ステップ名', '')).trim()
  };
}

function purchaseRowToValues_(r) {
  return [
    r.purchaseRequestId, r.requestDate, r.applicantEmail, r.applicantName, r.desiredDate,
    r.supplier, r.itemName, r.quantity, r.unitPrice, r.totalAmount, r.purpose,
    r.budgetCategory, r.paymentMethod, r.note, r.masterStatus || PURCHASE_MASTER_STATUS.REGISTERED,
    r.unregisteredMasterCount || 0, r.status, r.approverEmail, r.approvedAt,
    r.rejectReason, r.updatedAt, r.routeId || '', r.currentStep || 0, r.totalSteps || 0, r.currentStepName || ''
  ];
}

function writePurchaseRow_(purchase) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASES) || getSpreadsheet_().insertSheet(SHEET_PURCHASES);
  ensureHeaders_(sheet, PURCHASE_HEADERS);
  var all = readPurchaseRows_();
  var idx = -1;
  for (var i = 0; i < all.length; i++) {
    if (all[i].purchaseRequestId === purchase.purchaseRequestId) { idx = i; break; }
  }
  if (idx >= 0) all[idx] = purchase;
  else all.push(purchase);
  writeAllPurchaseRows_(sheet, all);
}

function writeAllPurchaseRows_(sheet, rows) {
  sheet.getRange(1, 1, 1, PURCHASE_HEADERS.length).setValues([PURCHASE_HEADERS]);
  if (!rows || rows.length === 0) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    return;
  }
  var values = rows.map(purchaseRowToValues_);
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  sheet.getRange(2, 1, values.length, PURCHASE_HEADERS.length).setValues(values);
}

function buildPurchaseRequest_(purchaseRequestId) {
  var rows = readPurchaseRows_(function(r) { return r.purchaseRequestId === purchaseRequestId; });
  return rows.length ? rows[0] : null;
}

function readPurchaseDetails_(purchaseRequestId) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASE_DETAILS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, PURCHASE_DETAIL_HEADERS.length).getValues();
  var id = String(purchaseRequestId || '').trim();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (id && String(data[i][0]).trim() !== id) continue;
    rows.push({
      purchaseRequestId: String(data[i][0] || '').trim(),
      lineNo: parseInt(data[i][1], 10) || (i + 1),
      maker: String(data[i][2] || '').trim(),
      modelNumber: String(data[i][3] || '').trim(),
      itemName: String(data[i][4] || '').trim(),
      quantity: normalizeAmount(data[i][5]),
      unitPrice: normalizeAmount(data[i][6]),
      amount: normalizeAmount(data[i][7]),
      note: String(data[i][8] || '').trim()
    });
  }
  rows.sort(function(a, b) { return a.lineNo - b.lineNo; });
  return rows;
}

function writePurchaseDetails_(purchaseRequestId, details) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASE_DETAILS) || getSpreadsheet_().insertSheet(SHEET_PURCHASE_DETAILS);
  ensureHeaders_(sheet, PURCHASE_DETAIL_HEADERS);
  var id = String(purchaseRequestId || '').trim();
  var otherRows = readPurchaseDetails_().filter(function(r) {
    return r.purchaseRequestId !== id;
  });
  var newRows = (details || []).map(function(d, index) {
    var quantity = normalizeAmount(d.quantity);
    var unitPrice = normalizeAmount(d.unitPrice);
    var amount = normalizeAmount(d.amount) || quantity * unitPrice;
    return {
      purchaseRequestId: id,
      lineNo: index + 1,
      maker: String(d.maker || '').trim(),
      modelNumber: String(d.modelNumber || '').trim(),
      itemName: String(d.itemName || '').trim(),
      quantity: quantity,
      unitPrice: unitPrice,
      amount: amount,
      note: String(d.note || '').trim()
    };
  });
  writeAllPurchaseDetails_(sheet, otherRows.concat(newRows));
}

function writeAllPurchaseDetails_(sheet, rows) {
  sheet.getRange(1, 1, 1, PURCHASE_DETAIL_HEADERS.length).setValues([PURCHASE_DETAIL_HEADERS]);
  if (!rows || rows.length === 0) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    return;
  }
  var values = rows.map(function(r) {
    return [
      r.purchaseRequestId, r.lineNo, r.maker, r.modelNumber, r.itemName,
      r.quantity, r.unitPrice, r.amount, r.note
    ];
  });
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  sheet.getRange(2, 1, values.length, PURCHASE_DETAIL_HEADERS.length).setValues(values);
}

function deletePurchaseDetails_(purchaseRequestId) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASE_DETAILS);
  if (!sheet) return;
  var id = String(purchaseRequestId || '').trim();
  var rows = readPurchaseDetails_().filter(function(r) {
    return r.purchaseRequestId !== id;
  });
  writeAllPurchaseDetails_(sheet, rows);
}

function buildPurchaseRequestWithDetails_(purchaseRequestId) {
  var purchase = buildPurchaseRequest_(purchaseRequestId);
  if (!purchase) return null;
  purchase.details = readPurchaseDetails_(purchaseRequestId);
  return purchase;
}

function getPurchaseDetailHistoryByModel_(modelNumber, excludePurchaseRequestId) {
  var key = String(modelNumber || '').normalize('NFKC').toLowerCase().replace(/\s/g, '').trim();
  if (!key) return { found: false, history: [] };

  var purchaseMap = {};
  readPurchaseRows_().forEach(function(p) {
    purchaseMap[p.purchaseRequestId] = p;
  });

  var excludeId = String(excludePurchaseRequestId || '').trim();
  var history = readPurchaseDetails_().filter(function(d) {
    if (excludeId && d.purchaseRequestId === excludeId) return false;
    return String(d.modelNumber || '').normalize('NFKC').toLowerCase().replace(/\s/g, '').trim() === key;
  }).map(function(d) {
    var p = purchaseMap[d.purchaseRequestId] || {};
    return {
      purchaseRequestId: d.purchaseRequestId,
      requestDate: p.requestDate || '',
      desiredDate: p.desiredDate || '',
      supplier: p.supplier || '',
      maker: d.maker || '',
      modelNumber: d.modelNumber || '',
      itemName: d.itemName || '',
      quantity: d.quantity || 0,
      unitPrice: d.unitPrice || 0,
      amount: d.amount || 0
    };
  });

  history.sort(function(a, b) {
    return (b.requestDate || '').localeCompare(a.requestDate || '');
  });

  return {
    found: history.length > 0,
    latest: history.length ? history[0] : null,
    history: history.slice(0, 5)
  };
}

function appendHistory_(purchaseRequestId, action, comment) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_HISTORY) || getSpreadsheet_().insertSheet(SHEET_HISTORY);
  ensureHeaders_(sheet, HISTORY_HEADERS);
  sheet.appendRow([purchaseRequestId, formatDateTime(new Date()), getCurrentUserEmail_(), action, comment || '']);
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
