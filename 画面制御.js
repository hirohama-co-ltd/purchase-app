function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  template.requestId = (e && e.parameter && e.parameter.requestId) || '';
  template.adminView = (e && e.parameter && e.parameter.admin) || '';
  template.portalUrl = PORTAL_URL;
  return template
    .evaluate()
    .setTitle('購買申請')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function filterPurchaseLists_(userEmail) {
  userEmail = String(userEmail || '').trim().toLowerCase();
  var all = readPurchaseRows_();
  var myPurchases = [];
  var pendingApprovals = [];
  for (var i = 0; i < all.length; i++) {
    var row = all[i];
    if (row.applicantEmail === userEmail) myPurchases.push(row);
    if (row.status === PURCHASE_STATUS.SUBMITTED && row.approverEmail === userEmail) pendingApprovals.push(row);
  }
  myPurchases.sort(function(a, b) { return (b.updatedAt || '').localeCompare(a.updatedAt || ''); });
  pendingApprovals.sort(function(a, b) { return (a.requestDate || '').localeCompare(b.requestDate || ''); });
  return { myPurchases: myPurchases, pendingApprovals: pendingApprovals };
}

function getBootstrapAppData() {
  var userEmail = getCurrentUserEmail_();
  var employee = findEmployeeByEmail(userEmail);
  return {
    userEmail: userEmail,
    employee: employee,
    isMasterAdmin: isPurchaseMasterAdmin_(userEmail),
    workflowRoutes: getAvailableWorkflowRoutes(),
    workflowLinked: isWorkflowLinked_(),
    statusLabels: PURCHASE_STATUS,
    purchaseMasters: getPurchaseMasterCandidates(),
    currencies: getPurchaseCurrencyOptions_()
  };
}

function getPurchaseListAppData() {
  var userEmail = getCurrentUserEmail_();
  return filterPurchaseLists_(userEmail);
}

function getSubmitReadinessApi() {
  var userEmail = getCurrentUserEmail_();
  var employee = findEmployeeByEmail(userEmail);
  return getSubmitReadiness_(userEmail, employee);
}

function previewPurchaseWorkflowRouteApi(routeId, applicantEmail) {
  return previewWorkflowRoute_(routeId, applicantEmail || getCurrentUserEmail_());
}

function lookupPurchaseDetailByModelApi(modelNumber, excludePurchaseRequestId) {
  return getPurchaseDetailHistoryByModel_(modelNumber, excludePurchaseRequestId);
}

function getPendingPurchaseMasterCandidatesApi() {
  var userEmail = getCurrentUserEmail_();
  if (!isPurchaseMasterAdmin_(userEmail)) {
    return { success: false, message: '未登録マスタ管理の権限がありません。', candidates: [] };
  }
  return { success: true, candidates: getOpenUnregisteredMasterCandidates_() };
}

function registerPendingPurchaseMasterCandidatesApi(updates) {
  var userEmail = getCurrentUserEmail_();
  if (!isPurchaseMasterAdmin_(userEmail)) {
    return { success: false, message: '未登録マスタ管理の権限がありません。' };
  }
  updates = updates || [];
  if (!updates.length) return { success: false, message: '登録対象がありません。' };
  updateUnregisteredMasterCandidateInputs_(updates);
  var ids = updates.map(function(u) { return String(u.candidateId || '').trim(); }).filter(function(id) { return !!id; });
  var message = registerPendingPurchaseMasters(ids);
  return {
    success: true,
    message: message,
    candidates: getOpenUnregisteredMasterCandidates_(),
    purchaseMasters: getPurchaseMasterCandidates()
  };
}
