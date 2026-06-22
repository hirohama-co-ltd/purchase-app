// ========================================
// 🔗 ワークフロー設定アプリ連携
// ========================================

var WF_SHEET_ROUTES = '申請経路マスタ';
var WF_SHEET_STEPS = '申請経路ステップ';
var WF_SHEET_APP_BINDING = 'アプリ別利用設定';
var WF_ROUTE_STATUS_COMPLETE = '完成';
var WF_STEP_TYPE_FIXED = '固定ユーザー';
var WF_STEP_TYPE_LEGACY_SUPERVISOR = '申請者の上長';
var WF_RESOLVE_ROLE_IN_ORG = '選択したロール内の指定した組織に所属するユーザーが承認';
var WF_RESOLVE_WF_ADMIN = '指定した組織のワークフロー管理者が承認';
var WF_RESOLVE_ORG_APPROVER = '組織内承認者が承認';
var WF_RESOLVE_ALL_IN_ORG = '組織に所属するすべてのユーザーが承認';
var WF_ADMIN_ROLE_NAME = 'ワークフロー管理者';
var WF_ORG_APPROVER_ROLES = ['課長', '事業所長', '部長', '工場長', '所長', 'マネージャー'];
var WF_APPROVAL_ROLE_APPROVER = '承認者';
var WF_APPROVAL_CONDITION_ONE_OR_MORE = '1人以上が承認';
var WF_REJECT_TARGET_PREVIOUS = 'ひとつ前の決裁者';
var WF_REJECT_TARGET_APPLICANT = '申請者';
var WF_REJECT_TARGET_ON_REJECT = '差戻し時に選択';
var WF_STEP_COL_COUNT = 11;

function isWorkflowLinked_() {
  return !!String(WORKFLOW_SS_ID || '').trim();
}

function openWorkflowSpreadsheet_() {
  if (!isWorkflowLinked_()) return null;
  try {
    return SpreadsheetApp.openById(WORKFLOW_SS_ID);
  } catch (e) {
    Logger.log('ワークフローブック open エラー: ' + e.message);
    return null;
  }
}

function getWorkflowAccessError_() {
  if (!isWorkflowLinked_()) return null;
  if (openWorkflowSpreadsheet_()) return null;
  return 'ワークフロー設定ブックを開けません。デプロイ実行ユーザーに閲覧権限を付与してください。';
}

function readWorkflowSheetRows_(sheetName, colCount, mapFn, cacheKey) {
  if (!isWorkflowLinked_()) return [];
  var cached = getCachedJson_(cacheKey);
  if (cached) return cached;

  var ss = openWorkflowSpreadsheet_();
  if (!ss) return [];
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) {
    putCachedJson_(cacheKey, []);
    return [];
  }
  var data = sheet.getRange(2, 1, sheet.getLastRow(), Math.max(colCount, sheet.getLastColumn())).getValues();
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    rows.push(mapFn(data[i]));
  }
  putCachedJson_(cacheKey, rows);
  return rows;
}

function loadWorkflowRoutes_() {
  return readWorkflowSheetRows_(WF_SHEET_ROUTES, 7, function(d) {
    return {
      routeId: String(d[0] || '').trim(),
      routeName: String(d[1] || '').trim(),
      status: String(d[2] || '').trim(),
      description: String(d[4] || '').trim()
    };
  }, 'wf_routes_' + WORKFLOW_SS_ID);
}

function loadWorkflowSteps_() {
  return readWorkflowSheetRows_(WF_SHEET_STEPS, WF_STEP_COL_COUNT, function(d) {
    return {
      routeId: String(d[0] || '').trim(),
      stepNo: parseInt(d[1], 10) || 0,
      stepName: String(d[2] || '').trim(),
      approverType: String(d[3] || '').trim(),
      approverEmail: String(d[4] || '').trim().toLowerCase(),
      approvalRole: String(d[5] || WF_APPROVAL_ROLE_APPROVER).trim() || WF_APPROVAL_ROLE_APPROVER,
      approvalCondition: String(d[6] || WF_APPROVAL_CONDITION_ONE_OR_MORE).trim() || WF_APPROVAL_CONDITION_ONE_OR_MORE,
      rejectTarget: String(d[7] || WF_REJECT_TARGET_APPLICANT).trim() || WF_REJECT_TARGET_APPLICANT,
      targetOffice: String(d[8] || '').trim(),
      targetDepartment: String(d[9] || '').trim(),
      targetRole: String(d[10] || '').trim()
    };
  }, 'wf_steps_' + WORKFLOW_SS_ID);
}

function loadWorkflowBindings_() {
  return readWorkflowSheetRows_(WF_SHEET_APP_BINDING, 5, function(d) {
    return {
      appCode: String(d[0] || '').trim(),
      routeId: String(d[2] || '').trim(),
      isDefault: String(d[3] || '').trim() === 'Y',
      active: String(d[4] || 'Y').trim() !== 'N'
    };
  }, 'wf_bindings_' + WORKFLOW_SS_ID);
}

function getAvailableWorkflowRoutes() {
  if (!isWorkflowLinked_() || getWorkflowAccessError_()) return [];
  var bindings = loadWorkflowBindings_().filter(function(b) {
    return b.appCode === APP_CODE && b.active;
  });
  var routeMap = {};
  loadWorkflowRoutes_().filter(function(r) { return r.status === WF_ROUTE_STATUS_COMPLETE; })
    .forEach(function(r) { routeMap[r.routeId] = r; });
  var steps = loadWorkflowSteps_();
  var list = [];
  bindings.forEach(function(b) {
    var r = routeMap[b.routeId];
    if (!r) return;
    list.push({
      routeId: r.routeId,
      routeName: r.routeName,
      description: r.description,
      stepCount: steps.filter(function(s) { return s.routeId === r.routeId; }).length,
      isDefault: b.isDefault
    });
  });
  list.sort(function(a, b) {
    if (a.isDefault !== b.isDefault) return a.isDefault ? -1 : 1;
    return a.routeName.localeCompare(b.routeName);
  });
  return list;
}

function getDefaultWorkflowRouteId_() {
  var routes = getAvailableWorkflowRoutes();
  for (var i = 0; i < routes.length; i++) {
    if (routes[i].isDefault) return routes[i].routeId;
  }
  return routes.length ? routes[0].routeId : '';
}

function validateWorkflowRoute_(routeId) {
  routeId = String(routeId || '').trim();
  if (!isWorkflowLinked_()) return null;
  if (!routeId) return '申請経路を選択してください。';
  var ok = loadWorkflowBindings_().some(function(b) {
    return b.appCode === APP_CODE && b.routeId === routeId && b.active;
  });
  if (!ok) return '選択した経路はこのアプリで利用できません。';
  var route = loadWorkflowRoutes_().filter(function(r) { return r.routeId === routeId; })[0];
  if (!route || route.status !== WF_ROUTE_STATUS_COMPLETE) return '経路が未完成または存在しません。';
  var steps = loadWorkflowSteps_().filter(function(s) { return s.routeId === routeId; });
  if (!steps.length) return '経路に承認ステップが設定されていません。';
  return null;
}

function wfNormalizeOrg_(value, fallback) {
  var v = String(value || '').trim();
  if (!v || v === '(申請者と同じ)') return String(fallback || '').trim();
  return v;
}

function wfFilterByOrg_(employees, office, department) {
  office = String(office || '').trim();
  department = String(department || '').trim();
  return employees.filter(function(e) {
    if (office && e.office !== office) return false;
    if (department && !employeeMatchesDepartmentFilter_(e, department)) return false;
    return true;
  });
}

function wfUniqueEmails_(employees) {
  var seen = {}, out = [];
  employees.forEach(function(e) {
    if (e.email && !seen[e.email]) { seen[e.email] = true; out.push(e.email); }
  });
  return out;
}

function wfIsOrgApproverRole_(role) {
  role = String(role || '').trim();
  if (!role) return false;
  if (role.indexOf('承認') !== -1) return true;
  for (var i = 0; i < WF_ORG_APPROVER_ROLES.length; i++) {
    if (role === WF_ORG_APPROVER_ROLES[i] || role.indexOf(WF_ORG_APPROVER_ROLES[i]) !== -1) return true;
  }
  return false;
}

function resolveWorkflowStepApprovers_(step, applicantEmail) {
  var applicant = findEmployeeByEmail(applicantEmail) || {};
  var office = wfNormalizeOrg_(step.targetOffice, applicant.office);
  var department = wfNormalizeOrg_(step.targetDepartment, applicant.department);
  var employees = loadEmployeesFromSheet();
  var type = String(step.approverType || '').trim();
  var emails = [];

  if (type === WF_STEP_TYPE_FIXED) {
    var fixed = String(step.approverEmail || '').trim().toLowerCase();
    if (fixed) emails = [fixed];
  } else if (type === WF_STEP_TYPE_LEGACY_SUPERVISOR) {
    return null;
  } else if (type === WF_RESOLVE_ROLE_IN_ORG) {
    var role = String(step.targetRole || '').trim();
    if (!role) return null;
    emails = wfUniqueEmails_(wfFilterByOrg_(employees, office, department).filter(function(e) {
      return employeeHasRole_(e, role);
    }));
  } else if (type === WF_RESOLVE_WF_ADMIN) {
    emails = wfUniqueEmails_(wfFilterByOrg_(employees, office, department).filter(function(e) {
      return employeeHasRole_(e, WF_ADMIN_ROLE_NAME);
    }));
  } else if (type === WF_RESOLVE_ORG_APPROVER) {
    var map = {};
    wfFilterByOrg_(employees, office, department).forEach(function(e) {
      if (employeeHasOrgApproverRole_(e) && e.email) map[e.email] = true;
    });
    emails = Object.keys(map);
  } else if (type === WF_RESOLVE_ALL_IN_ORG) {
    emails = wfUniqueEmails_(wfFilterByOrg_(employees, office, department));
  } else {
    return null;
  }

  if (!emails.length) return null;
  return { emails: emails, primaryEmail: emails[0], approverCount: emails.length };
}

function explainStepResolveFailure_(step) {
  var label = 'ステップ' + (step.stepNo || '?') + '「' + (step.stepName || '') + '」';
  return label + ': 承認者を解決できません。ワークフロー設定と社員マスタを確認してください。';
}

function resolveWorkflowSteps_(routeId, applicantEmail) {
  var steps = loadWorkflowSteps_()
    .filter(function(s) { return s.routeId === String(routeId || '').trim(); })
    .sort(function(a, b) { return a.stepNo - b.stepNo; });
  var resolved = [];
  for (var i = 0; i < steps.length; i++) {
    var match = resolveWorkflowStepApprovers_(steps[i], applicantEmail);
    if (!match || !match.primaryEmail) return [];
    resolved.push({
      stepNo: steps[i].stepNo,
      stepName: steps[i].stepName,
      approverEmail: match.primaryEmail,
      approverEmails: match.emails,
      approverCount: match.approverCount,
      approvalRole: steps[i].approvalRole || WF_APPROVAL_ROLE_APPROVER,
      approvalCondition: steps[i].approvalCondition || WF_APPROVAL_CONDITION_ONE_OR_MORE,
      rejectTarget: steps[i].rejectTarget || WF_REJECT_TARGET_APPLICANT
    });
  }
  return resolved;
}

function getWorkflowStep_(routeId, applicantEmail, stepNo) {
  var steps = resolveWorkflowSteps_(routeId, applicantEmail);
  stepNo = parseInt(stepNo, 10);
  for (var i = 0; i < steps.length; i++) {
    if (steps[i].stepNo === stepNo) return steps[i];
  }
  return null;
}

function getNextWorkflowStep_(routeId, applicantEmail, currentStepNo) {
  var steps = resolveWorkflowSteps_(routeId, applicantEmail);
  var nextNo = parseInt(currentStepNo, 10) + 1;
  for (var i = 0; i < steps.length; i++) {
    if (steps[i].stepNo === nextNo) return steps[i];
  }
  return null;
}

function resolveSubmitWorkflow_(routeId, applicantEmail) {
  if (!isWorkflowLinked_()) return { success: false, message: 'WORKFLOW_SS_ID が未設定です。' };
  var accessErr = getWorkflowAccessError_();
  if (accessErr) return { success: false, message: accessErr };
  var err = validateWorkflowRoute_(routeId);
  if (err) return { success: false, message: err };

  var rawSteps = loadWorkflowSteps_()
    .filter(function(s) { return s.routeId === routeId; })
    .sort(function(a, b) { return a.stepNo - b.stepNo; });
  if (!rawSteps.length) return { success: false, message: '経路に承認ステップが設定されていません。' };

  var resolved = [];
  for (var i = 0; i < rawSteps.length; i++) {
    var match = resolveWorkflowStepApprovers_(rawSteps[i], applicantEmail);
    if (!match || !match.primaryEmail) {
      return { success: false, message: explainStepResolveFailure_(rawSteps[i]) };
    }
    resolved.push({
      stepNo: rawSteps[i].stepNo,
      stepName: rawSteps[i].stepName,
      approvalRole: rawSteps[i].approvalRole || WF_APPROVAL_ROLE_APPROVER,
      approverEmail: match.primaryEmail
    });
  }

  var first = resolved[0];
  return {
    success: true,
    routeId: routeId,
    currentStep: first.stepNo,
    totalSteps: resolved.length,
    currentStepName: first.stepName,
    currentApprovalRole: first.approvalRole,
    approverEmail: first.approverEmail
  };
}

function enrichPurchaseWithWorkflowStep_(purchase) {
  if (!purchase || !isWorkflowLinked_() || !purchase.routeId || !purchase.currentStep) return purchase;
  var step = getWorkflowStep_(purchase.routeId, purchase.applicantEmail, purchase.currentStep);
  if (!step) return purchase;
  purchase.currentApprovalRole = step.approvalRole;
  purchase.currentApprovalCondition = step.approvalCondition;
  purchase.currentRejectTarget = step.rejectTarget;
  return purchase;
}

function resolveRejectRoute_(purchase, rejectTargetChoice) {
  var step = getWorkflowStep_(purchase.routeId, purchase.applicantEmail, purchase.currentStep);
  var target = step ? step.rejectTarget : WF_REJECT_TARGET_APPLICANT;
  if (target === WF_REJECT_TARGET_ON_REJECT) {
    target = String(rejectTargetChoice || '').trim() || WF_REJECT_TARGET_APPLICANT;
  }
  if (target === WF_REJECT_TARGET_PREVIOUS && purchase.currentStep > 1) {
    var prev = getWorkflowStep_(purchase.routeId, purchase.applicantEmail, purchase.currentStep - 1);
    if (prev) {
      return { mode: 'previous_step', stepNo: prev.stepNo, stepName: prev.stepName, approverEmail: prev.approverEmail };
    }
  }
  return { mode: 'applicant' };
}

function getSubmitReadiness_(userEmail, employee) {
  if (!userEmail) return { ok: false, message: 'ログインユーザーを取得できません。' };
  if (!employee) return { ok: false, message: '共通マスタの社員マスタに未登録です。' };
  var routeId = getDefaultWorkflowRouteId_();
  if (!routeId) {
    return { ok: false, message: '利用可能な申請経路がありません。ワークフロー設定で ' + APP_CODE + ' を紐づけてください。' };
  }
  var wf = resolveSubmitWorkflow_(routeId, userEmail);
  if (!wf.success) return { ok: false, message: wf.message };
  return { ok: true };
}

function previewWorkflowRoute_(routeId, applicantEmail) {
  applicantEmail = String(applicantEmail || '').trim();
  routeId = String(routeId || '').trim() || getDefaultWorkflowRouteId_();
  var err = validateWorkflowRoute_(routeId);
  if (err) return { success: false, message: err };
  var routeName = '';
  loadWorkflowRoutes_().some(function(r) {
    if (r.routeId === routeId) { routeName = r.routeName; return true; }
    return false;
  });
  var rawSteps = loadWorkflowSteps_().filter(function(s) { return s.routeId === routeId; }).sort(function(a, b) { return a.stepNo - b.stepNo; });
  var steps = [];
  for (var i = 0; i < rawSteps.length; i++) {
    var match = resolveWorkflowStepApprovers_(rawSteps[i], applicantEmail);
    if (!match || !match.primaryEmail) return { success: false, message: explainStepResolveFailure_(rawSteps[i]) };
    steps.push({
      stepNo: rawSteps[i].stepNo,
      stepName: rawSteps[i].stepName,
      approvalRole: rawSteps[i].approvalRole || WF_APPROVAL_ROLE_APPROVER,
      approverDisplay: displayNamesForEmails_(match.emails).join('、')
    });
  }
  return { success: true, routeId: routeId, routeName: routeName, steps: steps };
}
