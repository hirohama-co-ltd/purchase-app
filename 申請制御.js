// ========================================
// 📝 購買申請の作成・保存・提出
// ========================================

function validatePurchasePayload_(payload, isSubmit) {
  var errors = [];
  if (!payload.desiredDate) errors.push('希望納期を入力してください');
  if (!String(payload.supplier || '').trim()) errors.push('購入先を入力してください');
  if (!String(payload.purpose || '').trim()) errors.push('購買目的を入力してください');
  var details = normalizePurchaseDetails_(payload.details);
  if (!details.length) {
    errors.push('購買明細を1件以上入力してください');
  } else {
    details.forEach(function(d, index) {
      var label = '明細' + (index + 1) + ': ';
      if (!d.itemName) errors.push(label + '品名を入力してください');
      if (d.quantity <= 0) errors.push(label + '数量を入力してください');
      if (d.unitPrice <= 0) errors.push(label + '単価を入力してください');
    });
  }
  return errors;
}

function normalizePurchaseDetails_(details) {
  return (details || []).map(function(d) {
    var quantity = normalizeAmount(d.quantity);
    var unitPrice = normalizeAmount(d.unitPrice);
    return {
      maker: String(d.maker || '').trim(),
      modelNumber: String(d.modelNumber || '').trim(),
      itemName: String(d.itemName || '').trim(),
      quantity: quantity,
      unitPrice: unitPrice,
      amount: quantity * unitPrice,
      note: String(d.note || '').trim()
    };
  }).filter(function(d) {
    return d.maker || d.modelNumber || d.itemName || d.quantity || d.unitPrice || d.note;
  });
}

function summarizePurchaseItems_(details) {
  if (!details || !details.length) return '';
  if (details.length === 1) return details[0].itemName;
  return details[0].itemName + ' ほか' + (details.length - 1) + '件';
}

function savePurchaseRequest(payload, submit) {
  payload = payload || {};
  var userEmail = getCurrentUserEmail_();
  if (!userEmail) return { success: false, message: 'ログインユーザーを取得できません。' };

  var employee = findEmployeeByEmail(userEmail);
  var isSubmit = submit === true;
  if (isSubmit && !employee) return { success: false, message: '共通マスタの社員マスタに登録されていません。' };

  var errors = validatePurchasePayload_(payload, isSubmit);
  if (errors.length > 0) return { success: false, message: errors.join('\n') };

  var purchaseRequestId = String(payload.purchaseRequestId || '').trim() || generatePurchaseRequestId_();
  var existing = buildPurchaseRequest_(purchaseRequestId);
  if (existing) {
    if (existing.applicantEmail !== userEmail) return { success: false, message: '他のユーザーの申請は編集できません。' };
    if (existing.status !== PURCHASE_STATUS.DRAFT && existing.status !== PURCHASE_STATUS.REJECTED) {
      return { success: false, message: '現在のステータス（' + existing.status + '）では編集できません。' };
    }
  }

  var now = formatDateTime(new Date());
  var routeId = String(payload.routeId || '').trim() || (existing ? existing.routeId : '');
  if (isSubmit && isWorkflowLinked_() && !routeId) routeId = getDefaultWorkflowRouteId_();
  var wf = isSubmit
    ? resolveSubmitWorkflow_(routeId, userEmail)
    : { success: true, routeId: existing ? existing.routeId : routeId, currentStep: 0, totalSteps: 0, currentStepName: '', approverEmail: '' };
  if (isSubmit && !wf.success) return { success: false, message: wf.message };

  var details = normalizePurchaseDetails_(payload.details);
  var totalAmount = details.reduce(function(sum, d) { return sum + d.amount; }, 0);
  var firstDetail = details[0] || {};
  var purchase = {
    purchaseRequestId: purchaseRequestId,
    requestDate: existing ? existing.requestDate : normalizeDate(new Date()),
    applicantEmail: userEmail,
    applicantName: employee ? employee.name : (payload.applicantName || userEmail.split('@')[0]),
    desiredDate: normalizeDate(payload.desiredDate),
    supplier: String(payload.supplier || '').trim(),
    itemName: summarizePurchaseItems_(details),
    quantity: details.reduce(function(sum, d) { return sum + d.quantity; }, 0),
    unitPrice: firstDetail.unitPrice || 0,
    totalAmount: totalAmount,
    purpose: String(payload.purpose || '').trim(),
    budgetCategory: String(payload.budgetCategory || '').trim(),
    paymentMethod: String(payload.paymentMethod || '').trim(),
    note: String(payload.note || ''),
    status: isSubmit ? PURCHASE_STATUS.SUBMITTED : PURCHASE_STATUS.DRAFT,
    approverEmail: isSubmit ? wf.approverEmail : (existing ? existing.approverEmail : ''),
    approvedAt: isSubmit ? '' : (existing ? existing.approvedAt : ''),
    rejectReason: isSubmit ? '' : (existing ? existing.rejectReason : ''),
    updatedAt: now,
    routeId: isSubmit ? wf.routeId : (existing ? existing.routeId : routeId),
    currentStep: isSubmit ? wf.currentStep : (existing ? existing.currentStep : 0),
    totalSteps: isSubmit ? wf.totalSteps : (existing ? existing.totalSteps : 0),
    currentStepName: isSubmit ? wf.currentStepName : (existing ? existing.currentStepName : '')
  };

  writePurchaseRow_(purchase);
  writePurchaseDetails_(purchaseRequestId, details);
  appendHistory_(purchaseRequestId, isSubmit ? '申請' : '下書き保存', '');
  return {
    success: true,
    purchaseRequestId: purchaseRequestId,
    message: isSubmit ? '購買申請を提出しました。' : '下書きを保存しました。',
    purchase: buildPurchaseRequest_(purchaseRequestId)
  };
}

function deletePurchaseRequest(purchaseRequestId) {
  var userEmail = getCurrentUserEmail_();
  var purchase = buildPurchaseRequest_(purchaseRequestId);
  if (!purchase) return { success: false, message: '申請が見つかりません。' };
  if (purchase.applicantEmail !== userEmail) return { success: false, message: '削除権限がありません。' };
  if (purchase.status !== PURCHASE_STATUS.DRAFT) return { success: false, message: '下書きのみ削除できます。' };
  var all = readPurchaseRows_(function(r) { return r.purchaseRequestId !== purchaseRequestId; });
  writeAllPurchaseRows_(getSpreadsheet_().getSheetByName(SHEET_PURCHASES), all);
  deletePurchaseDetails_(purchaseRequestId);
  appendHistory_(purchaseRequestId, '削除', '');
  return { success: true, message: '下書きを削除しました。' };
}

function withdrawPurchaseRequest(purchaseRequestId) {
  var userEmail = getCurrentUserEmail_();
  var purchase = buildPurchaseRequest_(purchaseRequestId);
  if (!purchase) return { success: false, message: '申請が見つかりません。' };
  if (purchase.applicantEmail !== userEmail) return { success: false, message: '取り下げ権限がありません。' };
  if (purchase.status !== PURCHASE_STATUS.SUBMITTED) return { success: false, message: '申請中の申請のみ取り下げできます。' };
  purchase.status = PURCHASE_STATUS.WITHDRAWN;
  purchase.approverEmail = '';
  purchase.approvedAt = '';
  purchase.rejectReason = '';
  purchase.currentStep = 0;
  purchase.totalSteps = 0;
  purchase.currentStepName = '';
  purchase.updatedAt = formatDateTime(new Date());
  writePurchaseRow_(purchase);
  appendHistory_(purchaseRequestId, '取り下げ', '');
  return { success: true, message: '申請を取り下げました。' };
}

function getPurchaseRequestDetail(purchaseRequestId) {
  var purchase = buildPurchaseRequestWithDetails_(purchaseRequestId);
  if (!purchase) return { success: false, message: '申請が見つかりません。' };
  var userEmail = getCurrentUserEmail_();
  var canView = purchase.applicantEmail === userEmail || purchase.approverEmail === userEmail;
  if (!canView) return { success: false, message: '閲覧権限がありません。' };
  return { success: true, purchase: enrichPurchaseWithWorkflowStep_(purchase) };
}
