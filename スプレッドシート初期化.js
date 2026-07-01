/**
 * 出張申請アプリ用スプレッドシートの一括初期化
 */

var TRIP_SAMPLES = [
  ['BT-20250601-0001', '2025-06-01', 'yamada@example.com', '山田太郎', '2025-06-10 09:00', true, '2025-06-11 18:00', false,
    '大阪市', '取引先訪問', '新幹線', 'ホテルニューオータニ大阪', 12000, 50000, '', '承認済', '未精算', '', 'manager@example.com', '', '', '',
    'WF-TRIP-STANDARD', 1, 1, '直属上長承認']
];

function initializeSpreadsheet(options) {
  options = options || {};
  var addSamples = options.addSamples !== false;
  var forceHeaders = options.forceHeaders === true;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logs = [];

  var specs = [
    { name: SHEET_TRIPS, headers: TRIP_HEADERS, tabColor: '#bfdbfe', samples: addSamples ? TRIP_SAMPLES : null },
    { name: SHEET_HISTORY, headers: HISTORY_HEADERS, tabColor: '#e9d5ff' }
  ];

  specs.forEach(function(spec, index) {
    logs.push(ensureSheetWithHeaders_(ss, spec, forceHeaders));
    var sheet = ss.getSheetByName(spec.name);
    if (sheet) {
      try { sheet.setTabColor(spec.tabColor); sheet.setPosition(index + 1); } catch (e) { /* ignore */ }
    }
  });

  writeSetupGuideSheet_(ss, forceHeaders);
  clearMasterCaches_();

  var created = logs.filter(function(l) { return l.created; }).map(function(l) { return l.name; });
  var msg = 'スプレッドシート初期化が完了しました。\n\n';
  if (created.length) msg += '【新規作成】\n・' + created.join('\n・') + '\n\n';
  msg += '【次の作業】\n';
  msg += '1. このスプレッドシートの ID を出張旅費精算アプリの 設定.js → TRIP_APP_SS_ID に設定\n';
  msg += '2. ワークフロー設定ブック ID を 設定.js → WORKFLOW_SS_ID に設定\n';
  msg += '3. Webアプリとしてデプロイ\n';
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

  if (spec.samples && (created || forceHeaders) && spec.samples.length > 0) {
    var padded = spec.samples.map(function(row) {
      var copy = row.slice();
      while (copy.length < spec.headers.length) copy.push('');
      return copy.slice(0, spec.headers.length);
    });
    if (created || sheet.getLastRow() <= 1) {
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      sheet.getRange(2, 1, padded.length, spec.headers.length).setValues(padded);
    }
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
    ['出張申請アプリ｜セットアップ手順'],
    [],
    ['手順', '内容'],
    ['1', 'メニュー「出張申請」→「全シート＋ヘッダーを一括作成」'],
    ['2', 'Webアプリとしてデプロイ'],
    ['3', 'スプレッドシート URL から ID をコピー → 出張旅費精算 設定.js の TRIP_APP_SS_ID に設定'],
    ['4', 'ワークフロー設定ブック ID → 設定.js の WORKFLOW_SS_ID に設定'],
    ['5', '出張旅費精算アプリも clasp push → 再デプロイ'],
    [],
    ['ID の形式', ''],
    ['出張申請ID', 'BT-yyyyMMdd-xxxx（例: BT-20250610-1234）'],
    ['精算ID', 'TR-yyyyMMdd-xxxx（出張旅費精算アプリが付与）'],
    [],
    ['ステータス（出張申請）', ''],
    ['下書き', '編集中'],
    ['申請中', '承認待ち'],
    ['承認済', '出張可 → 旅費精算アプリで精算開始可'],
    ['差戻し', '修正して再申請'],
    ['取り下げ', '申請者が申請を撤回'],
    ['取消', '承認後の取消（精算前のみ）'],
    [],
    ['精算状況', ''],
    ['未精算', '旅費精算未作成'],
    ['精算中', '旅費精算を提出済み'],
    ['精算完了', '旅費精算が精算完了']
  ];
  sheet.getRange(1, 1, guide.length, 2).setValues(normalizeGuideRowsTo2Cols_(guide));
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(12);
  sheet.setColumnWidths(1, 1, 200);
  sheet.setColumnWidths(2, 1, 420);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('出張申請')
    .addItem('全シート＋ヘッダーを一括作成', 'menuInitializeSpreadsheet')
    .addToUi();
}

function menuInitializeSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('初期化', 'シートとヘッダーを作成します。実行しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert('完了', initializeSpreadsheet({ addSamples: false, forceHeaders: false }), ui.ButtonSet.OK);
}
