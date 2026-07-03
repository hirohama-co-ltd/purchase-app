/**
 * 購買申請アプリ用スプレッドシートの一括初期化
 */

function initializeSpreadsheet(options) {
  options = options || {};
  var forceHeaders = options.forceHeaders === true;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logs = [];
  var specs = [
    { name: SHEET_PURCHASES, headers: PURCHASE_HEADERS, tabColor: '#fde68a' },
    { name: SHEET_PURCHASE_DETAILS, headers: PURCHASE_DETAIL_HEADERS, tabColor: '#fed7aa' },
    { name: SHEET_SUPPLIER_MASTER, headers: PURCHASE_MASTER_HEADERS, tabColor: '#bbf7d0' },
    { name: SHEET_MAKER_MASTER, headers: PURCHASE_MASTER_HEADERS, tabColor: '#bfdbfe' },
    { name: SHEET_UNREGISTERED_MASTER_CANDIDATES, headers: UNREGISTERED_MASTER_HEADERS, tabColor: '#fecaca' },
    { name: SHEET_HISTORY, headers: HISTORY_HEADERS, tabColor: '#e9d5ff' },
    { name: SHEET_RECURRING, headers: RECURRING_HEADERS, tabColor: '#ddd6fe' },
    { name: SHEET_RECURRING_DETAILS, headers: RECURRING_DETAIL_HEADERS, tabColor: '#e9d5ff' }
  ];

  specs.forEach(function(spec, index) {
    logs.push(ensureSheetWithHeaders_(ss, spec, forceHeaders));
    var sheet = ss.getSheetByName(spec.name);
    if (sheet) {
      try { sheet.setTabColor(spec.tabColor); sheet.setPosition(index + 1); } catch (e) { /* ignore */ }
    }
  });

  writeSetupGuideSheet_(ss, forceHeaders);
  clearPurchaseNameMasterCaches_();
  clearMasterCaches_();

  var created = logs.filter(function(l) { return l.created; }).map(function(l) { return l.name; });
  var msg = 'スプレッドシート初期化が完了しました。\n\n';
  if (created.length) msg += '【新規作成】\n・' + created.join('\n・') + '\n\n';
  msg += '【次の作業】\n';
  msg += '1. ワークフロー設定で APP_CODE「' + APP_CODE + '」を経路に紐づけ\n';
  msg += '2. Webアプリとしてデプロイ\n';
  msg += '3. 申請ポータルの「ポータル連携」にこのアプリを登録（dataType: purchase）\n';
  msg += '4. Apps Script エディタで installRecurringPurchaseTrigger を実行（定期申請）\n';
  return msg;
}

function ensureSheetWithHeaders_(ss, spec, forceHeaders) {
  var sheet = ss.getSheetByName(spec.name);
  var created = false;
  if (!sheet) { sheet = ss.insertSheet(spec.name); created = true; }
  var headerNeedsUpdate = created || forceHeaders || sheet.getLastRow() === 0 || !headersMatch_(
    sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), spec.headers.length)).getValues()[0],
    spec.headers
  );
  if (headerNeedsUpdate) {
    sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    sheet.getRange(1, 1, 1, spec.headers.length).setFontWeight('bold').setBackground('#e2e8f0');
  }
  sheet.setFrozenRows(1);
  return { name: spec.name, created: created };
}

function headersMatch_(row, expected) {
  for (var i = 0; i < expected.length; i++) {
    if (String(row[i] || '').trim() !== expected[i]) return false;
  }
  return true;
}

/** 2列ガイド用：各行を必ず2列に揃える（setValues エラー防止） */
function normalizeGuideRowsTo2Cols_(rows) {
  return rows.map(function(row) {
    row = row || [];
    return [
      row[0] != null ? String(row[0]) : '',
      row[1] != null ? String(row[1]) : ''
    ];
  });
}

function writeSetupGuideSheet_(ss, forceRewrite) {
  var name = 'セットアップ手順';
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  else if (!forceRewrite && sheet.getLastRow() > 3) return;

  sheet.clear();
  var guide = [
    ['購買申請アプリ｜セットアップ手順'],
    [],
    ['手順', '内容'],
    ['1', 'メニュー「購買申請」→「全シート＋ヘッダーを一括作成」'],
    ['2', 'ワークフロー設定で APP_CODE「PURCHASE_REQUEST」をアプリ別利用に登録'],
    ['3', 'Webアプリとしてデプロイ'],
    ['4', '申請ポータルのポータル連携で dataType「purchase」として登録'],
    ['5', '購入先マスタ・メーカーマスタに候補を登録（コード / 表示名 / 別名 / 有効）'],
    ['6', '未登録マスタ候補は、正式表示名 / コード / 別名を入力してメニュー「未登録マスタを正式登録」を実行'],
    ['7', 'Apps Script エディタで installRecurringPurchaseTrigger を実行（毎月自動申請）'],
    [],
    ['マスタ別名の例', '表示名: 三菱電機 / 別名: 三菱,MITSUBISHI,ミツビシ'],
    [],
    ['ID の形式', 'PR-yyyyMMdd-xxxx'],
    [],
    ['ステータス', '下書き / 申請中 / 承認済 / 差戻し / 取り下げ / 取消']
  ];
  sheet.getRange(1, 1, guide.length, 2).setValues(normalizeGuideRowsTo2Cols_(guide));
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(12);
  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 1, 520);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('購買申請')
    .addItem('全シート＋ヘッダーを一括作成', 'menuInitializeSpreadsheet')
    .addSeparator()
    .addItem('未登録マスタを正式登録', 'menuRegisterPendingPurchaseMasters')
    .addSeparator()
    .addItem('為替レート取得を確認', 'menuAuthorizeExchangeRateAccess')
    .addItem('定期申請トリガーを設定', 'menuInstallRecurringPurchaseTrigger')
    .addToUi();
}

function menuAuthorizeExchangeRateAccess() {
  var ui = SpreadsheetApp.getUi();
  var result = authorizeExchangeRateAccess();
  if (result.success) {
    ui.alert('確認完了',
      '為替レートを取得できました。\n例: 1 USD = ' + Number(result.rate || 0).toLocaleString('ja-JP') + ' 円',
      ui.ButtonSet.OK);
    return;
  }
  ui.alert('為替レート取得',
    result.message || '為替レートの取得に失敗しました。',
    ui.ButtonSet.OK);
}

function menuInstallRecurringPurchaseTrigger() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('定期申請トリガー', '毎日 7:00 JST に定期購買の自動申請チェックを実行するトリガーを設定します。よろしいですか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }
  ui.alert('完了', installRecurringPurchaseTrigger(), ui.ButtonSet.OK);
}

function menuInitializeSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('初期化', 'シートとヘッダーを作成します。実行しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert('完了', initializeSpreadsheet({ forceHeaders: false }), ui.ButtonSet.OK);
}

function menuRegisterPendingPurchaseMasters() {
  var ui = SpreadsheetApp.getUi();
  var message = '未登録マスタ候補のうち、正式表示名が入力されている行を正式登録します。\n実行しますか？';
  if (ui.alert('未登録マスタを正式登録', message, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert('完了', registerPendingPurchaseMasters(), ui.ButtonSet.OK);
}
