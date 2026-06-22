// ========================================
// 🏷️ 購入先・メーカーマスタ制御
// ========================================

var PURCHASE_MASTER_HEADERS = ['コード', '表示名', '別名', '有効'];

function normalizeMasterKey_(value) {
  return String(value || '')
    .normalize('NFKC')
    .toLowerCase()
    .replace(/\s/g, '')
    .trim();
}

function loadPurchaseNameMaster_(sheetName) {
  var cacheKey = 'purchase_name_master_' + sheetName;
  var cached = getCachedJson_(cacheKey);
  if (cached) return cached;

  var sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    putCachedJson_(cacheKey, []);
    return [];
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, PURCHASE_MASTER_HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var displayName = String(data[i][1] || '').trim();
    if (!displayName) continue;
    if (String(data[i][3] || 'Y').trim() === 'N') continue;
    rows.push({
      code: String(data[i][0] || '').trim(),
      displayName: displayName,
      aliases: parseMultiValues_(data[i][2])
    });
  }
  putCachedJson_(cacheKey, rows);
  return rows;
}

function resolvePurchaseMasterName_(input, sheetName, label, required) {
  var raw = String(input || '').trim();
  if (!raw) {
    return required ? { success: false, message: label + 'を入力してください。' } : { success: true, value: '' };
  }

  var rows = loadPurchaseNameMaster_(sheetName);
  if (!rows.length) {
    return { success: true, value: raw };
  }

  var key = normalizeMasterKey_(raw);
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (normalizeMasterKey_(row.displayName) === key || normalizeMasterKey_(row.code) === key) {
      return { success: true, value: row.displayName };
    }
    for (var j = 0; j < row.aliases.length; j++) {
      if (normalizeMasterKey_(row.aliases[j]) === key) {
        return { success: true, value: row.displayName };
      }
    }
  }
  return { success: true, value: raw, unregistered: true, inputName: raw, label: label };
}

function resolveSupplierName_(input) {
  return resolvePurchaseMasterName_(input, SHEET_SUPPLIER_MASTER, '購入先', true);
}

function resolveMakerName_(input) {
  return resolvePurchaseMasterName_(input, SHEET_MAKER_MASTER, 'メーカー', false);
}

function getPurchaseMasterCandidates() {
  return {
    suppliers: loadPurchaseNameMaster_(SHEET_SUPPLIER_MASTER).map(function(r) { return r.displayName; }),
    makers: loadPurchaseNameMaster_(SHEET_MAKER_MASTER).map(function(r) { return r.displayName; })
  };
}

function clearPurchaseNameMasterCaches_() {
  try {
    CacheService.getScriptCache().remove('purchase_name_master_' + SHEET_SUPPLIER_MASTER);
    CacheService.getScriptCache().remove('purchase_name_master_' + SHEET_MAKER_MASTER);
  } catch (e) { /* ignore */ }
}

function getUnregisteredMasterCandidateSheet_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_UNREGISTERED_MASTER_CANDIDATES) ||
    getSpreadsheet_().insertSheet(SHEET_UNREGISTERED_MASTER_CANDIDATES);
  ensureHeaders_(sheet, UNREGISTERED_MASTER_HEADERS);
  return sheet;
}

function makeUnregisteredCandidateId_(type, inputName, purchaseRequestId, lineNo) {
  return [
    type,
    normalizeMasterKey_(inputName),
    String(purchaseRequestId || '').trim(),
    String(lineNo || 0)
  ].join('|');
}

function readUnregisteredMasterCandidates_(filterFn) {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_UNREGISTERED_MASTER_CANDIDATES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, UNREGISTERED_MASTER_HEADERS.length).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var row = {
      candidateId: String(data[i][0] || '').trim(),
      type: String(data[i][1] || '').trim(),
      inputName: String(data[i][2] || '').trim(),
      officialName: String(data[i][3] || '').trim(),
      code: String(data[i][4] || '').trim(),
      aliases: String(data[i][5] || '').trim(),
      status: String(data[i][6] || '未登録').trim() || '未登録',
      purchaseRequestId: String(data[i][7] || '').trim(),
      lineNo: parseInt(data[i][8], 10) || 0,
      updatedAt: formatDateTime(data[i][9]),
      registeredAt: formatDateTime(data[i][10]),
      processedBy: String(data[i][11] || '').trim()
    };
    if (!filterFn || filterFn(row)) rows.push(row);
  }
  return rows;
}

function writeAllUnregisteredMasterCandidates_(rows) {
  var sheet = getUnregisteredMasterCandidateSheet_();
  sheet.getRange(1, 1, 1, UNREGISTERED_MASTER_HEADERS.length).setValues([UNREGISTERED_MASTER_HEADERS]);
  if (!rows || rows.length === 0) {
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
    return;
  }
  var values = rows.map(function(r) {
    return [
      r.candidateId, r.type, r.inputName, r.officialName || '', r.code || '', r.aliases || '',
      r.status || '未登録', r.purchaseRequestId || '', r.lineNo || 0, r.updatedAt || '',
      r.registeredAt || '', r.processedBy || ''
    ];
  });
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  sheet.getRange(2, 1, values.length, UNREGISTERED_MASTER_HEADERS.length).setValues(values);
}

function recordUnregisteredMasterCandidates_(purchaseRequestId, pendingItems) {
  var now = formatDateTime(new Date());
  var rows = readUnregisteredMasterCandidates_();
  var map = {};
  rows.forEach(function(r) { map[r.candidateId] = r; });
  var currentIds = {};

  (pendingItems || []).forEach(function(item) {
    var id = makeUnregisteredCandidateId_(item.type, item.inputName, purchaseRequestId, item.lineNo);
    currentIds[id] = true;
    var existing = map[id] || {};
    map[id] = {
      candidateId: id,
      type: item.type,
      inputName: item.inputName,
      officialName: existing.officialName || '',
      code: existing.code || '',
      aliases: existing.aliases || '',
      status: existing.status === '登録済' ? '登録済' : '未登録',
      purchaseRequestId: purchaseRequestId,
      lineNo: item.lineNo || 0,
      updatedAt: now,
      registeredAt: existing.registeredAt || '',
      processedBy: existing.processedBy || ''
    };
  });

  Object.keys(map).forEach(function(id) {
    var row = map[id];
    if (row.purchaseRequestId === purchaseRequestId && row.status === '未登録' && !currentIds[id]) {
      row.status = '不要';
      row.updatedAt = now;
    }
  });

  writeAllUnregisteredMasterCandidates_(Object.keys(map).map(function(k) { return map[k]; }));
}

function countOpenUnregisteredCandidates_(purchaseRequestId) {
  var id = String(purchaseRequestId || '').trim();
  return readUnregisteredMasterCandidates_(function(r) {
    return r.purchaseRequestId === id && r.status === '未登録';
  }).length;
}

function updatePurchaseMasterStatus_(purchaseRequestId) {
  var purchase = buildPurchaseRequest_(purchaseRequestId);
  if (!purchase) return;
  var count = countOpenUnregisteredCandidates_(purchaseRequestId);
  purchase.unregisteredMasterCount = count;
  purchase.masterStatus = count > 0 ? PURCHASE_MASTER_STATUS.PENDING : PURCHASE_MASTER_STATUS.REGISTERED;
  purchase.updatedAt = formatDateTime(new Date());
  writePurchaseRow_(purchase);
}

function isPurchaseMasterAdmin_(email) {
  return employeeHasRole_(findEmployeeByEmail(email || getCurrentUserEmail_()), WF_ADMIN_ROLE_NAME);
}

function getOpenUnregisteredMasterCandidates_() {
  return readUnregisteredMasterCandidates_(function(r) {
    return r.status === '未登録';
  }).sort(function(a, b) {
    return (b.updatedAt || '').localeCompare(a.updatedAt || '');
  });
}

function updateUnregisteredMasterCandidateInputs_(updates) {
  var now = formatDateTime(new Date());
  var updateMap = {};
  (updates || []).forEach(function(u) {
    var id = String(u.candidateId || '').trim();
    if (!id) return;
    updateMap[id] = {
      officialName: String(u.officialName || '').trim(),
      code: String(u.code || '').trim(),
      aliases: String(u.aliases || '').trim()
    };
  });
  if (!Object.keys(updateMap).length) return [];

  var rows = readUnregisteredMasterCandidates_();
  rows.forEach(function(r) {
    var update = updateMap[r.candidateId];
    if (!update || r.status !== '未登録') return;
    r.officialName = update.officialName;
    r.code = update.code;
    r.aliases = update.aliases;
    r.updatedAt = now;
  });
  writeAllUnregisteredMasterCandidates_(rows);
  return rows;
}

function appendNameMasterRow_(sheetName, code, officialName, aliases) {
  var sheet = getSpreadsheet_().getSheetByName(sheetName) || getSpreadsheet_().insertSheet(sheetName);
  ensureHeaders_(sheet, PURCHASE_MASTER_HEADERS);
  var rows = loadPurchaseNameMaster_(sheetName);
  var resolved = resolvePurchaseMasterName_(officialName, sheetName, sheetName, false);
  if (rows.length && resolved.success && !resolved.unregistered) return;
  sheet.appendRow([code || officialName, officialName, aliases || '', 'Y']);
}

function applyOfficialNameToPurchaseHistory_(candidate) {
  var officialName = candidate.officialName || candidate.inputName;
  if (candidate.type === '購入先') {
    var all = readPurchaseRows_();
    all.forEach(function(p) {
      if (p.purchaseRequestId === candidate.purchaseRequestId && normalizeMasterKey_(p.supplier) === normalizeMasterKey_(candidate.inputName)) {
        p.supplier = officialName;
        p.updatedAt = formatDateTime(new Date());
      }
    });
    writeAllPurchaseRows_(getSpreadsheet_().getSheetByName(SHEET_PURCHASES), all);
  } else if (candidate.type === 'メーカー') {
    var detailSheet = getSpreadsheet_().getSheetByName(SHEET_PURCHASE_DETAILS);
    if (!detailSheet) return;
    var details = readPurchaseDetails_();
    details.forEach(function(d) {
      if (d.purchaseRequestId === candidate.purchaseRequestId &&
          (!candidate.lineNo || d.lineNo === candidate.lineNo) &&
          normalizeMasterKey_(d.maker) === normalizeMasterKey_(candidate.inputName)) {
        d.maker = officialName;
      }
    });
    writeAllPurchaseDetails_(detailSheet, details);
  }
}

function registerPendingPurchaseMasters(candidateIds) {
  var targetMap = {};
  (candidateIds || []).forEach(function(id) {
    id = String(id || '').trim();
    if (id) targetMap[id] = true;
  });
  var hasTarget = Object.keys(targetMap).length > 0;
  var rows = readUnregisteredMasterCandidates_();
  var now = formatDateTime(new Date());
  var userEmail = getCurrentUserEmail_();
  var processed = 0;
  var messages = [];

  rows.forEach(function(r) {
    if (r.status !== '未登録') return;
    if (hasTarget && !targetMap[r.candidateId]) return;
    if (!r.officialName) {
      messages.push(r.type + '「' + r.inputName + '」: 正式表示名が未入力です。');
      return;
    }
    var sheetName = r.type === '購入先' ? SHEET_SUPPLIER_MASTER : SHEET_MAKER_MASTER;
    appendNameMasterRow_(sheetName, r.code, r.officialName, r.aliases);
    applyOfficialNameToPurchaseHistory_(r);
    r.status = '登録済';
    r.registeredAt = now;
    r.updatedAt = now;
    r.processedBy = userEmail;
    processed++;
  });

  writeAllUnregisteredMasterCandidates_(rows);
  clearPurchaseNameMasterCaches_();

  var affected = {};
  rows.forEach(function(r) {
    if (r.purchaseRequestId) affected[r.purchaseRequestId] = true;
  });
  Object.keys(affected).forEach(updatePurchaseMasterStatus_);

  return '正式登録処理が完了しました。\n登録: ' + processed + ' 件' +
    (messages.length ? '\n\n【未処理】\n' + messages.join('\n') : '');
}
