// ========================================
// ✅ 購買申請の承認・差戻し
// ========================================

function approvePurchaseRequest(purchaseRequestId, comment) {
  return updatePurchaseStatus_(purchaseRequestId, PURCHASE_STATUS.APPROVED, comment || '承認しました');
}

function rejectPurchaseRequest(purchaseRequestId, reason, rejectTargetChoice) {
  reason = String(reason || '').trim();
  if (!reason) return { success: false, message: '差戻し理由を入力してください。' };
  return updatePurchaseStatus_(purchaseRequestId, PURCHASE_STATUS.REJECTED, reason, rejectTargetChoice);
}

function updatePurchaseStatus_(purchaseRequestId, newStatus, comment, rejectTargetChoice) {
  var userEmail = getCurrentUserEmail_();
  var purchase = buildPurchaseRequest_(purchaseRequestId);
  if (!purchase) return { success: false, message: '申請が見つかりません。' };
  if (purchase.status !== PURCHASE_STATUS.SUBMITTED) return { success: false, message: '申請中のデータのみ承認・差戻しできます。' };
  if (purchase.approverEmail !== userEmail) return { success: false, message: '承認権限がありません。' };

  if (newStatus === PURCHASE_STATUS.APPROVED) {
    var nextStep = getNextWorkflowStep_(purchase.routeId, purchase.applicantEmail, purchase.currentStep);
    if (nextStep) {
      purchase.currentStep = nextStep.stepNo;
      purchase.currentStepName = nextStep.stepName;
      purchase.approverEmail = nextStep.approverEmail;
      purchase.updatedAt = formatDateTime(new Date());
      writePurchaseRow_(purchase);
      appendHistory_(purchaseRequestId, '承認（' + (purchase.currentStep - 1) + '/' + purchase.totalSteps + '）', comment || '承認しました');
      return {
        success: true,
        message: '承認しました。次の承認者（' + nextStep.stepName + '）に回りました。',
        purchase: buildPurchaseRequest_(purchaseRequestId)
      };
    }
  }

  if (newStatus === PURCHASE_STATUS.REJECTED) {
    var route = resolveRejectRoute_(purchase, rejectTargetChoice);
    if (route.mode === 'previous_step') {
      purchase.currentStep = route.stepNo;
      purchase.currentStepName = route.stepName;
      purchase.approverEmail = route.approverEmail;
      purchase.rejectReason = comment;
      purchase.updatedAt = formatDateTime(new Date());
      writePurchaseRow_(purchase);
      appendHistory_(purchaseRequestId, '差戻し（前ステップへ）', comment || '');
      return {
        success: true,
        message: '差戻しました。前の承認者（' + route.stepName + '）に戻しました。',
        purchase: enrichPurchaseWithWorkflowStep_(buildPurchaseRequest_(purchaseRequestId))
      };
    }
  }

  purchase.status = newStatus;
  purchase.updatedAt = formatDateTime(new Date());
  if (newStatus === PURCHASE_STATUS.APPROVED) {
    purchase.approvedAt = purchase.updatedAt;
    purchase.rejectReason = '';
  }
  if (newStatus === PURCHASE_STATUS.REJECTED) {
    purchase.rejectReason = comment;
    purchase.approvedAt = '';
    purchase.currentStep = 0;
    purchase.currentStepName = '';
  }
  writePurchaseRow_(purchase);

  var actionLabel = newStatus === PURCHASE_STATUS.APPROVED ? '承認' : '差戻し';
  appendHistory_(purchaseRequestId, actionLabel, comment || '');
  return { success: true, message: actionLabel + 'しました。', purchase: buildPurchaseRequest_(purchaseRequestId) };
}
