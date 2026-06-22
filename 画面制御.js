function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  template.requestId = (e && e.parameter && e.parameter.requestId) || '';
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
    workflowRoutes: getAvailableWorkflowRoutes(),
    workflowLinked: isWorkflowLinked_(),
    statusLabels: PURCHASE_STATUS
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
