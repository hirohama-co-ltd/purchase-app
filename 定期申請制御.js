// ========================================
// 🔁 購買申請 — 定期購買（毎月自動申請）
// ========================================

function normalizeRecurringSchedule_(schedule) {
  schedule = schedule || {};
  var runDay = parseInt(schedule.runDayOfMonth, 10);
  if (isNaN(runDay) || runDay < 1) runDay = 1;
  if (runDay > 28) runDay = 28;
  var offset = parseInt(schedule.desiredDateOffsetDays, 10);
  if (isNaN(offset) || offset < 0) offset = DEFAULT_RECURRING_DESIRED_OFFSET_DAYS;
  return {
    runDayOfMonth: runDay,
    desiredDateOffsetDays: offset
  };
}

function buildRecurringSummaryLabel_(template) {
  template = template || {};
  var parts = [];
  if (template.supplier) parts.push(template.supplier);
  if (template.itemName) parts.push(template.itemName);
  return parts.join(' / ') || template.recurringId || '定期購買';
}

function enrichRecurringTemplateForList_(template) {
  var details = readRecurringDetails_(template.recurringId);
  return {
    recurringId: template.recurringId,
    status: template.status,
    applicantEmail: template.applicantEmail,
    applicantName: template.applicantName,
    runDayOfMonth: template.runDayOfMonth,
    desiredDateOffsetDays: template.desiredDateOffsetDays,
    supplier: template.supplier,
    itemName: summarizePurchaseItems_(recurringDetailsToPurchaseDetails_(details)),
    purpose: template.purpose,
    budgetCategory: template.budgetCategory,
    paymentMethod: template.paymentMethod,
    currency: template.currency || DEFAULT_PURCHASE_CURRENCY,
    note: template.note,
    sourcePurchaseRequestId: template.sourcePurchaseRequestId,
    lastRunDate: template.lastRunDate,
    lastGeneratedPurchaseRequestId: template.lastGeneratedPurchaseRequestId,
    registeredAt: template.registeredAt,
    updatedAt: template.updatedAt,
    summaryLabel: buildRecurringSummaryLabel_({
      supplier: template.supplier,
      itemName: summarizePurchaseItems_(recurringDetailsToPurchaseDetails_(details))
    })
  };
}

function listRecurringTemplatesForUser_(userEmail) {
  userEmail = String(userEmail || '').trim().toLowerCase();
  return readRecurringRows_(function(row) {
    return row.applicantEmail === userEmail;
  }).map(enrichRecurringTemplateForList_).sort(function(a, b) {
    return (b.updatedAt || '').localeCompare(a.updatedAt || '');
  });
}

function getRecurringTemplateForUser_(recurringId, userEmail) {
  recurringId = String(recurringId || '').trim();
  userEmail = String(userEmail || '').trim().toLowerCase();
  var template = buildRecurringTemplateWithDetails_(recurringId);
  if (!template) return { success: false, message: '定期購買設定が見つかりません。' };
  if (template.applicantEmail !== userEmail) {
    return { success: false, message: 'この定期購買設定を操作する権限がありません。' };
  }
  return { success: true, template: template };
}

function saveRecurringTemplateFromPayload_(payload, schedule, userEmail) {
  payload = payload || {};
  userEmail = String(userEmail || getCurrentUserEmail_() || '').trim().toLowerCase();
  if (!userEmail) return { success: false, message: 'ログインユーザーを取得できません。' };

  var employee = findEmployeeByEmail(userEmail);
  if (!employee) return { success: false, message: '共通マスタの社員マスタに登録されていません。' };

  var nameResult = normalizePurchasePayloadNames_(payload);
  if (nameResult.errors.length > 0) return { success: false, message: nameResult.errors.join('\n') };

  var errors = validatePurchasePayload_(payload, true);
  if (errors.length > 0) return { success: false, message: errors.join('\n') };

  schedule = normalizeRecurringSchedule_(schedule);
  var details = normalizePurchaseDetails_(payload.details);
  var now = formatDateTime(new Date());
  var routeId = String(payload.routeId || '').trim();
  if (isWorkflowLinked_() && !routeId) routeId = getDefaultWorkflowRouteId_();

  var template = {
    recurringId: generateRecurringId_(),
    status: RECURRING_STATUS.ACTIVE,
    applicantEmail: userEmail,
    applicantName: employee.name,
    runDayOfMonth: schedule.runDayOfMonth,
    desiredDateOffsetDays: schedule.desiredDateOffsetDays,
    routeId: routeId,
    supplier: String(payload.supplier || '').trim(),
    purpose: String(payload.purpose || '').trim(),
    budgetCategory: String(payload.budgetCategory || '').trim(),
    paymentMethod: String(payload.paymentMethod || '').trim(),
    currency: normalizeCurrencyCode_(payload.currency),
    note: String(payload.note || '').trim(),
    sourcePurchaseRequestId: String(payload.sourcePurchaseRequestId || '').trim(),
    lastRunDate: '',
    lastGeneratedPurchaseRequestId: '',
    registeredAt: now,
    updatedAt: now
  };

  writeRecurringRow_(template);
  writeRecurringDetails_(template.recurringId, details);

  return {
    success: true,
    message: '定期購買を登録しました（毎月' + template.runDayOfMonth + '日に自動申請）。',
    template: enrichRecurringTemplateForList_(template)
  };
}

function saveRecurringTemplateFromPurchase_(purchaseRequestId, schedule, userEmail) {
  purchaseRequestId = String(purchaseRequestId || '').trim();
  userEmail = String(userEmail || getCurrentUserEmail_() || '').trim().toLowerCase();
  if (!purchaseRequestId) return { success: false, message: '購買申請IDが指定されていません。' };

  var purchase = buildPurchaseRequestWithDetails_(purchaseRequestId);
  if (!purchase) return { success: false, message: '購買申請が見つかりません。' };
  if (purchase.applicantEmail !== userEmail) {
    return { success: false, message: '自分の申請のみ定期化できます。' };
  }
  if (purchase.status !== PURCHASE_STATUS.APPROVED) {
    return { success: false, message: '承認済みの申請のみ定期化できます。' };
  }

  var payload = {
    routeId: purchase.routeId,
    supplier: purchase.supplier,
    purpose: purchase.purpose,
    budgetCategory: purchase.budgetCategory,
    paymentMethod: purchase.paymentMethod,
    note: purchase.note,
    currency: purchase.currency || DEFAULT_PURCHASE_CURRENCY,
    desiredDate: purchase.desiredDate,
    sourcePurchaseRequestId: purchase.purchaseRequestId,
    details: purchase.details || []
  };
  return saveRecurringTemplateFromPayload_(payload, schedule, userEmail);
}

function stopRecurringTemplate_(recurringId, userEmail) {
  var access = getRecurringTemplateForUser_(recurringId, userEmail);
  if (!access.success) return access;
  var template = access.template;
  template.status = RECURRING_STATUS.STOPPED;
  template.updatedAt = formatDateTime(new Date());
  writeRecurringRow_(template);
  return {
    success: true,
    message: '定期購買を停止しました。',
    template: enrichRecurringTemplateForList_(template)
  };
}

function resumeRecurringTemplate_(recurringId, userEmail) {
  var access = getRecurringTemplateForUser_(recurringId, userEmail);
  if (!access.success) return access;
  var template = access.template;
  template.status = RECURRING_STATUS.ACTIVE;
  template.updatedAt = formatDateTime(new Date());
  writeRecurringRow_(template);
  return {
    success: true,
    message: '定期購買を再開しました。',
    template: enrichRecurringTemplateForList_(template)
  };
}

function wasRecurringRunThisMonth_(template, yearMonth) {
  yearMonth = String(yearMonth || '').trim();
  if (!template.lastRunDate || !yearMonth) return false;
  return String(template.lastRunDate).substring(0, 7) === yearMonth.substring(0, 7);
}

function calculateRecurringDesiredDate_(runDate, offsetDays) {
  var d = new Date(runDate.getTime());
  d.setDate(d.getDate() + (parseInt(offsetDays, 10) || DEFAULT_RECURRING_DESIRED_OFFSET_DAYS));
  return normalizeDate(d);
}

function buildPurchasePayloadFromRecurring_(template, runDate) {
  var details = readRecurringDetails_(template.recurringId);
  return {
    routeId: template.routeId,
    desiredDate: calculateRecurringDesiredDate_(runDate, template.desiredDateOffsetDays),
    supplier: template.supplier,
    purpose: template.purpose,
    budgetCategory: template.budgetCategory,
    paymentMethod: template.paymentMethod,
    currency: template.currency || DEFAULT_PURCHASE_CURRENCY,
    note: template.note,
    details: recurringDetailsToPurchaseDetails_(details)
  };
}

function createPurchaseFromRecurring_(template, runDate) {
  var payload = buildPurchasePayloadFromRecurring_(template, runDate);
  var result = savePurchaseRequestCore_(payload, true, {
    actorEmail: template.applicantEmail,
    recurringId: template.recurringId,
    autoGenerated: true,
    skipOwnershipCheck: true,
    historyAction: '定期申請（自動）',
    historyComment: '定期ID: ' + template.recurringId
  });
  if (!result.success) return result;

  template.lastRunDate = normalizeDate(runDate);
  template.lastGeneratedPurchaseRequestId = result.purchaseRequestId;
  template.updatedAt = formatDateTime(new Date());
  writeRecurringRow_(template);

  notifyRecurringPurchaseCreated_(template, result.purchaseRequestId);
  return result;
}

function processRecurringTemplate_(template, runDate) {
  var tz = Session.getScriptTimeZone();
  var yearMonth = Utilities.formatDate(runDate, tz, 'yyyy-MM');
  if (wasRecurringRunThisMonth_(template, yearMonth)) {
    return { success: true, skipped: true, message: '今月は実行済み: ' + template.recurringId };
  }
  try {
    return createPurchaseFromRecurring_(template, runDate);
  } catch (e) {
    Logger.log('processRecurringTemplate_(' + template.recurringId + '): ' + e.message);
    return { success: false, message: e.message || String(e) };
  }
}

function runRecurringPurchaseJobs() {
  var runDate = new Date();
  var tz = Session.getScriptTimeZone();
  var dayOfMonth = parseInt(Utilities.formatDate(runDate, tz, 'd'), 10);
  var yearMonth = Utilities.formatDate(runDate, tz, 'yyyy-MM');
  var templates = readRecurringRows_(function(row) {
    return row.status === RECURRING_STATUS.ACTIVE && row.runDayOfMonth === dayOfMonth;
  });

  var summary = {
    date: normalizeDate(runDate),
    dayOfMonth: dayOfMonth,
    yearMonth: yearMonth,
    processed: 0,
    skipped: 0,
    failed: 0,
    messages: []
  };

  templates.forEach(function(template) {
    if (wasRecurringRunThisMonth_(template, yearMonth)) {
      summary.skipped++;
      return;
    }
    var result = processRecurringTemplate_(template, runDate);
    if (result.skipped) summary.skipped++;
    else if (result.success) summary.processed++;
    else {
      summary.failed++;
      summary.messages.push(template.recurringId + ': ' + (result.message || '失敗'));
    }
  });

  Logger.log('runRecurringPurchaseJobs: ' + JSON.stringify(summary));
  return summary;
}

function installRecurringPurchaseTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runRecurringPurchaseJobs') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('runRecurringPurchaseJobs')
    .timeBased()
    .atHour(RECURRING_RUN_HOUR_JST)
    .everyDays(1)
    .inTimezone(Session.getScriptTimeZone())
    .create();
  return '定期申請トリガーを設定しました（毎日 ' + RECURRING_RUN_HOUR_JST + ':00 JST に runRecurringPurchaseJobs を実行）。';
}

function notifyRecurringPurchaseCreated_(template, purchaseRequestId) {
  template = template || {};
  purchaseRequestId = String(purchaseRequestId || '').trim();
  var email = String(template.applicantEmail || '').trim().toLowerCase();
  if (!email || !purchaseRequestId) return;

  var baseUrl = getPurchaseWebAppUrl_();
  var url = baseUrl + (baseUrl.indexOf('?') >= 0 ? '&' : '?') + 'requestId=' + encodeURIComponent(purchaseRequestId);
  var subject = '[購買申請] 定期購買の自動申請を作成しました（' + purchaseRequestId + '）';
  var body = [
    template.applicantName || email,
    '',
    '定期購買設定（' + template.recurringId + '）に基づき、購買申請を自動提出しました。',
    '',
    '購買申請ID: ' + purchaseRequestId,
    '購入先: ' + (template.supplier || '-'),
    '毎月実行日: ' + template.runDayOfMonth + '日',
    '',
    '申請内容を確認する:',
    url,
    '',
    '支払いが不要になった場合は、購買申請アプリの「定期申請」タブから停止してください。'
  ].join('\n');

  try {
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log('notifyRecurringPurchaseCreated_: ' + e.message);
  }
}

function getRecurringTemplatesApi() {
  var userEmail = getCurrentUserEmail_();
  if (!userEmail) return { success: false, message: 'ログインユーザーを取得できません。', templates: [] };
  return {
    success: true,
    templates: listRecurringTemplatesForUser_(userEmail),
    webAppUrl: getPurchaseWebAppUrl_()
  };
}

function saveRecurringTemplateFromFormApi(payload, schedule) {
  try {
    return saveRecurringTemplateFromPayload_(payload, schedule, getCurrentUserEmail_());
  } catch (e) {
    Logger.log('saveRecurringTemplateFromFormApi: ' + e.message);
    return { success: false, message: e.message || String(e) };
  }
}

function saveRecurringTemplateFromPurchaseApi(purchaseRequestId, schedule) {
  try {
    return saveRecurringTemplateFromPurchase_(purchaseRequestId, schedule, getCurrentUserEmail_());
  } catch (e) {
    Logger.log('saveRecurringTemplateFromPurchaseApi: ' + e.message);
    return { success: false, message: e.message || String(e) };
  }
}

function stopRecurringTemplateApi(recurringId) {
  try {
    return stopRecurringTemplate_(recurringId, getCurrentUserEmail_());
  } catch (e) {
    Logger.log('stopRecurringTemplateApi: ' + e.message);
    return { success: false, message: e.message || String(e) };
  }
}

function resumeRecurringTemplateApi(recurringId) {
  try {
    return resumeRecurringTemplate_(recurringId, getCurrentUserEmail_());
  } catch (e) {
    Logger.log('resumeRecurringTemplateApi: ' + e.message);
    return { success: false, message: e.message || String(e) };
  }
}
