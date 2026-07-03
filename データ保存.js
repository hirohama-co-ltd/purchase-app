// ========================================
// 💾 購買申請データ読み書き
// ========================================

var PURCHASE_HEADERS = [
  '購買申請ID', '申請日', '申請者Email', '申請者名', '希望納期',
  '購入先', '品名', '数量', '単価', '合計金額', '購買目的', '予算区分', '支払方法', '備考',
  'マスタ登録状態', '未登録マスタ件数', 'ステータス', '承認者Email', '承認日時', '差戻し理由', '更新日時',
  '経路ID', '現在ステップ', '総ステップ数', '現在ステップ名',
  '定期ID', '自動生成', '通貨', '為替レート', '予想円換算'
];

var RECURRING_HEADERS = [
  '定期ID', '状態', '申請者Email', '申請者名', '毎月実行日', '希望納期オフセット日数',
  '経路ID', '購入先', '購買目的', '予算区分', '支払方法', '通貨', '備考',
  '元申請ID', '最終実行日', '最終生成申請ID', '登録日時', '更新日時'
];

var RECURRING_DETAIL_HEADERS = [
  '定期ID', '行No', 'メーカー', '品番', '品名', '数量', '単価', '金額', '備考'
];

var HISTORY_HEADERS = ['購買申請ID', '操作日時', '操作者Email', '操作', 'コメント'];
var PURCHASE_DETAIL_HEADERS = ['購買申請ID', '行No', 'メーカー', '品番', '品名', '数量', '単価', '金額', '備考'];
var UNREGISTERED_MASTER_HEADERS = ['候補ID', '種別', '入力名', '正式表示名', 'コード', '別名', '状態', '関連申請ID', '関連行No', '更新日時', '登録日時', '処理者Email'];

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  if (PURCHASE_SS_ID) return SpreadsheetApp.openById(PURCHASE_SS_ID);
  throw new Error('購買申請スプレッドシートを開けません。');
}

function buildHeaderColumnMap_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var row = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];
  var map = {};
  for (var i = 0; i < row.length; i++) {
    var name = String(row[i] || '').trim();
    if (name) map[name] = i;
  }
  return map;
}

function writeHeaderCells_(sheet, headers, rowIndex) {
  headers.forEach(function(name, index) {
    sheet.getRange(rowIndex, index + 1).setValue(name);
  });
}

function ensureHeaderRowByValues_(sheet, headers) {
  if (!sheet || !headers || !headers.length) return;
  var colMap = buildHeaderColumnMap_(sheet);
  var hasHeader = colMap.hasOwnProperty(headers[0]);
  if (!hasHeader) {
    if (sheet.getLastRow() >= 1) {
      sheet.insertRowBefore(1);
    }
    writeHeaderCells_(sheet, headers, 1);
    sheet.setFrozenRows(1);
    colMap = buildHeaderColumnMap_(sheet);
  }
  var nextCol = Math.max(sheet.getLastColumn(), 1) + 1;
  headers.forEach(function(name) {
    if (colMap.hasOwnProperty(name)) return;
    sheet.getRange(1, nextCol).setValue(name);
    colMap[name] = nextCol - 1;
    nextCol++;
  });
  sheet.setFrozenRows(1);
}

function getPurchaseColumnMap_(sheet) {
  var colMap = buildHeaderColumnMap_(sheet);
  var map = {};
  PURCHASE_HEADERS.forEach(function(name) {
    if (colMap.hasOwnProperty(name)) map[name] = colMap[name];
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
    currency: normalizeCurrencyCode_(purchaseCell_(data, colMap, '通貨', DEFAULT_PURCHASE_CURRENCY)),
    exchangeRate: parseFloat(purchaseCell_(data, colMap, '為替レート', 1)) || 1,
    estimatedJpyAmount: normalizeAmount(purchaseCell_(data, colMap, '予想円換算', purchaseCell_(data, colMap, '合計金額', 0))),
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
    currentStepName: String(purchaseCell_(data, colMap, '現在ステップ名', '')).trim(),
    recurringId: String(purchaseCell_(data, colMap, '定期ID', '')).trim(),
    autoGenerated: String(purchaseCell_(data, colMap, '自動生成', '')).trim()
  };
}

function purchaseRowToValues_(r) {
  return [
    r.purchaseRequestId, r.requestDate, r.applicantEmail, r.applicantName, r.desiredDate,
    r.supplier, r.itemName, r.quantity, r.unitPrice, r.totalAmount, r.purpose,
    r.budgetCategory, r.paymentMethod, r.note, r.masterStatus || PURCHASE_MASTER_STATUS.REGISTERED,
    r.unregisteredMasterCount || 0, r.status, r.approverEmail, r.approvedAt,
    r.rejectReason, r.updatedAt, r.routeId || '', r.currentStep || 0, r.totalSteps || 0, r.currentStepName || '',
    r.recurringId || '', r.autoGenerated || '',
    r.currency || DEFAULT_PURCHASE_CURRENCY, r.exchangeRate || 1, r.estimatedJpyAmount || r.totalAmount || 0
  ];
}

function purchaseToFieldMap_(purchase) {
  return {
    '購買申請ID': purchase.purchaseRequestId,
    '申請日': purchase.requestDate,
    '申請者Email': purchase.applicantEmail,
    '申請者名': purchase.applicantName,
    '希望納期': purchase.desiredDate,
    '購入先': purchase.supplier,
    '品名': purchase.itemName,
    '数量': purchase.quantity,
    '単価': purchase.unitPrice,
    '合計金額': purchase.totalAmount,
    '購買目的': purchase.purpose,
    '予算区分': purchase.budgetCategory,
    '支払方法': purchase.paymentMethod,
    '備考': purchase.note,
    'マスタ登録状態': purchase.masterStatus || PURCHASE_MASTER_STATUS.REGISTERED,
    '未登録マスタ件数': purchase.unregisteredMasterCount || 0,
    'ステータス': purchase.status,
    '承認者Email': purchase.approverEmail,
    '承認日時': purchase.approvedAt,
    '差戻し理由': purchase.rejectReason,
    '更新日時': purchase.updatedAt,
    '経路ID': purchase.routeId || '',
    '現在ステップ': purchase.currentStep || 0,
    '総ステップ数': purchase.totalSteps || 0,
    '現在ステップ名': purchase.currentStepName || '',
    '定期ID': purchase.recurringId || '',
    '自動生成': purchase.autoGenerated || '',
    '通貨': purchase.currency || DEFAULT_PURCHASE_CURRENCY,
    '為替レート': purchase.exchangeRate || 1,
    '予想円換算': purchase.estimatedJpyAmount || purchase.totalAmount || 0
  };
}

function writeRowByHeaderMap_(sheet, rowIndex, colMap, fieldMap) {
  Object.keys(fieldMap).forEach(function(name) {
    if (!colMap.hasOwnProperty(name)) return;
    sheet.getRange(rowIndex, colMap[name] + 1).setValue(fieldMap[name]);
  });
}

function normalizeRowWidth_(row, width) {
  row = (row || []).slice();
  while (row.length < width) row.push('');
  if (row.length > width) row.length = width;
  return row;
}

function writeDataRows_(sheet, startRow, rows, colCount) {
  if (!rows || rows.length === 0) return;
  var values = rows.map(function(r) {
    return normalizeRowWidth_(r, colCount);
  });
  var endRow = startRow + values.length - 1;
  sheet.getRange(startRow, 1, endRow, colCount).setValues(values);
}

function findSheetRowByFirstColumnId_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow, 1).getValues();
  var target = String(id || '').trim();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === target) return i + 2;
  }
  return -1;
}

function writePurchaseRow_(purchase) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASES) || getSpreadsheet_().insertSheet(SHEET_PURCHASES);
  ensureHeaderRowByValues_(sheet, PURCHASE_HEADERS);
  var colMap = getPurchaseColumnMap_(sheet);
  var rowIndex = findSheetRowByFirstColumnId_(sheet, purchase.purchaseRequestId);
  if (rowIndex <= 0) rowIndex = Math.max(sheet.getLastRow(), 1) + 1;
  writeRowByHeaderMap_(sheet, rowIndex, colMap, purchaseToFieldMap_(purchase));
}

function writeAllPurchaseRows_(sheet, rows) {
  sheet.getRange(1, 1, 1, PURCHASE_HEADERS.length).setValues([PURCHASE_HEADERS]);
  if (!rows || rows.length === 0) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    return;
  }
  var values = rows.map(function(r) {
    return normalizeRowWidth_(purchaseRowToValues_(r), PURCHASE_HEADERS.length);
  });
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  writeDataRows_(sheet, 2, values, PURCHASE_HEADERS.length);
}

function buildPurchaseRequest_(purchaseRequestId) {
  var rows = readPurchaseRows_(function(r) { return r.purchaseRequestId === purchaseRequestId; });
  return rows.length ? rows[0] : null;
}

function readPurchaseDetails_(purchaseRequestId) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASE_DETAILS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow(), PURCHASE_DETAIL_HEADERS.length).getValues();
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

function deletePurchaseDetailRows_(purchaseRequestId) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASE_DETAILS);
  if (!sheet || sheet.getLastRow() < 2) return;
  var id = String(purchaseRequestId || '').trim();
  for (var row = sheet.getLastRow(); row >= 2; row--) {
    if (String(sheet.getRange(row, 1).getValue() || '').trim() === id) {
      sheet.deleteRow(row);
    }
  }
}

function appendPurchaseDetailRows_(purchaseRequestId, details) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASE_DETAILS) ||
    getSpreadsheet_().insertSheet(SHEET_PURCHASE_DETAILS);
  ensureHeaderRowByValues_(sheet, PURCHASE_DETAIL_HEADERS);
  (details || []).forEach(function(d, index) {
    sheet.appendRow([
      purchaseRequestId, index + 1, String(d.maker || '').trim(), String(d.modelNumber || '').trim(),
      String(d.itemName || '').trim(), d.quantity, d.unitPrice, d.amount, String(d.note || '').trim()
    ]);
  });
}

function writePurchaseDetails_(purchaseRequestId, details) {
  deletePurchaseDetailRows_(purchaseRequestId);
  appendPurchaseDetailRows_(purchaseRequestId, details);
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
  writeDataRows_(sheet, 2, values, PURCHASE_DETAIL_HEADERS.length);
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

function appendHistory_(purchaseRequestId, action, comment, actorEmail) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_HISTORY) || getSpreadsheet_().insertSheet(SHEET_HISTORY);
  ensureHeaders_(sheet, HISTORY_HEADERS);
  var email = String(actorEmail || getCurrentUserEmail_() || '').trim().toLowerCase();
  sheet.appendRow([purchaseRequestId, formatDateTime(new Date()), email, action, comment || '']);
}

function getRecurringColumnMap_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), RECURRING_HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || '').trim();
  });
  var map = {};
  RECURRING_HEADERS.forEach(function(name) {
    var idx = headers.indexOf(name);
    if (idx >= 0) map[name] = idx;
  });
  return map;
}

function recurringCell_(data, colMap, name, defaultValue) {
  if (!colMap.hasOwnProperty(name)) return defaultValue;
  return data[colMap[name]];
}

function mapRecurringRow_(data, colMap) {
  return {
    recurringId: String(recurringCell_(data, colMap, '定期ID', '')).trim(),
    status: String(recurringCell_(data, colMap, '状態', RECURRING_STATUS.ACTIVE)).trim() || RECURRING_STATUS.ACTIVE,
    applicantEmail: String(recurringCell_(data, colMap, '申請者Email', '')).trim().toLowerCase(),
    applicantName: String(recurringCell_(data, colMap, '申請者名', '')).trim(),
    runDayOfMonth: parseInt(recurringCell_(data, colMap, '毎月実行日', 1), 10) || 1,
    desiredDateOffsetDays: parseInt(recurringCell_(data, colMap, '希望納期オフセット日数', DEFAULT_RECURRING_DESIRED_OFFSET_DAYS), 10) ||
      DEFAULT_RECURRING_DESIRED_OFFSET_DAYS,
    routeId: String(recurringCell_(data, colMap, '経路ID', '')).trim(),
    supplier: String(recurringCell_(data, colMap, '購入先', '')).trim(),
    purpose: String(recurringCell_(data, colMap, '購買目的', '')).trim(),
    budgetCategory: String(recurringCell_(data, colMap, '予算区分', '')).trim(),
    paymentMethod: String(recurringCell_(data, colMap, '支払方法', '')).trim(),
    currency: normalizeCurrencyCode_(recurringCell_(data, colMap, '通貨', DEFAULT_PURCHASE_CURRENCY)),
    note: String(recurringCell_(data, colMap, '備考', '')).trim(),
    sourcePurchaseRequestId: String(recurringCell_(data, colMap, '元申請ID', '')).trim(),
    lastRunDate: normalizeDate(recurringCell_(data, colMap, '最終実行日', '')),
    lastGeneratedPurchaseRequestId: String(recurringCell_(data, colMap, '最終生成申請ID', '')).trim(),
    registeredAt: formatDateTime(recurringCell_(data, colMap, '登録日時', '')),
    updatedAt: formatDateTime(recurringCell_(data, colMap, '更新日時', ''))
  };
}

function recurringRowToValues_(row) {
  return [
    row.recurringId, row.status, row.applicantEmail, row.applicantName, row.runDayOfMonth,
    row.desiredDateOffsetDays, row.routeId || '', row.supplier, row.purpose, row.budgetCategory,
    row.paymentMethod, row.currency || DEFAULT_PURCHASE_CURRENCY, row.note, row.sourcePurchaseRequestId || '', row.lastRunDate || '',
    row.lastGeneratedPurchaseRequestId || '', row.registeredAt || '', row.updatedAt || ''
  ];
}

function readRecurringRows_(filterFn) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RECURRING);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var colMap = getRecurringColumnMap_(sheet);
  var lastCol = Math.max(sheet.getLastColumn(), RECURRING_HEADERS.length);
  var data = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var id = String(recurringCell_(data[i], colMap, '定期ID', '')).trim();
    if (!id) continue;
    var row = mapRecurringRow_(data[i], colMap);
    if (!filterFn || filterFn(row)) rows.push(row);
  }
  return rows;
}

function buildRecurringTemplate_(recurringId) {
  var rows = readRecurringRows_(function(r) { return r.recurringId === recurringId; });
  return rows.length ? rows[0] : null;
}

function writeRecurringRow_(row) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RECURRING) || getSpreadsheet_().insertSheet(SHEET_RECURRING);
  ensureHeaders_(sheet, RECURRING_HEADERS);
  var values = normalizeRowWidth_(recurringRowToValues_(row), RECURRING_HEADERS.length);
  var rowIndex = findSheetRowByFirstColumnId_(sheet, row.recurringId);
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, rowIndex, RECURRING_HEADERS.length).setValues([values]);
    return;
  }
  var nextRow = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(nextRow, 1, nextRow, RECURRING_HEADERS.length).setValues([values]);
}

function readRecurringDetails_(recurringId) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RECURRING_DETAILS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow(), RECURRING_DETAIL_HEADERS.length).getValues();
  var id = String(recurringId || '').trim();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (id && String(data[i][0]).trim() !== id) continue;
    rows.push({
      recurringId: String(data[i][0] || '').trim(),
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

function writeRecurringDetails_(recurringId, details) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RECURRING_DETAILS) ||
    getSpreadsheet_().insertSheet(SHEET_RECURRING_DETAILS);
  ensureHeaders_(sheet, RECURRING_DETAIL_HEADERS);
  var id = String(recurringId || '').trim();
  var otherRows = readAllRecurringDetails_().filter(function(r) {
    return r.recurringId !== id;
  });
  var newRows = (details || []).map(function(d, index) {
    var quantity = normalizeAmount(d.quantity);
    var unitPrice = normalizeAmount(d.unitPrice);
    var amount = normalizeAmount(d.amount) || quantity * unitPrice;
    return {
      recurringId: id,
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
  writeAllRecurringDetails_(sheet, otherRows.concat(newRows));
}

function readAllRecurringDetails_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_RECURRING_DETAILS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow(), RECURRING_DETAIL_HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push({
      recurringId: String(data[i][0] || '').trim(),
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
  return rows;
}

function writeAllRecurringDetails_(sheet, rows) {
  sheet.getRange(1, 1, 1, RECURRING_DETAIL_HEADERS.length).setValues([RECURRING_DETAIL_HEADERS]);
  if (!rows || rows.length === 0) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    return;
  }
  var values = rows.map(function(r) {
    return [
      r.recurringId, r.lineNo, r.maker, r.modelNumber, r.itemName,
      r.quantity, r.unitPrice, r.amount, r.note
    ];
  });
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  writeDataRows_(sheet, 2, values, RECURRING_DETAIL_HEADERS.length);
}

function buildRecurringTemplateWithDetails_(recurringId) {
  var template = buildRecurringTemplate_(recurringId);
  if (!template) return null;
  template.details = readRecurringDetails_(recurringId);
  return template;
}

function recurringDetailsToPurchaseDetails_(details) {
  return (details || []).map(function(d) {
    return {
      maker: d.maker,
      modelNumber: d.modelNumber,
      itemName: d.itemName,
      quantity: d.quantity,
      unitPrice: d.unitPrice,
      amount: d.amount,
      note: d.note
    };
  });
}

function ensureHeaders_(sheet, headers) {
  ensureHeaderRowByValues_(sheet, headers);
  var headerWidth = Math.max(sheet.getLastColumn(), headers.length);
  sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setBackground('#e2e8f0');
}
