/**
 * 経理入力システム v2.0 - シート初期化
 * GASエディタから initializeSystem() を手動実行してください
 */

function initializeSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'システム初期化',
    '全シートを作成し、初期データを投入します。\n既存のシートは上書きされません。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // スプレッドシートIDを保存
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  let created = 0;
  let skipped = 0;

  // === マスタシート群 ===
  Object.values(SHEET_NAMES).forEach(name => {
    if (ss.getSheetByName(name)) {
      skipped++;
      return;
    }
    const sheet = ss.insertSheet(name);
    const headers = SHEET_HEADERS[name];
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      formatHeaderRow(sheet, headers.length);
    }
    created++;
  });

  // === 取引データシート（銀行・クレカ） ===
  const transactionSheets = [
    SHEET_NAMES.GMO, SHEET_NAMES.SBI, SHEET_NAMES.MIZUHO,
    SHEET_NAMES.CC_JCB, SHEET_NAMES.CC_VISA
  ];
  transactionSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, TRANSACTION_HEADERS.length).setValues([TRANSACTION_HEADERS]);
      formatHeaderRow(sheet, TRANSACTION_HEADERS.length);
    }
  });

  // === デフォルトデータ投入 ===
  insertDefaultSettings(ss);
  insertDefaultAttendanceSettings(ss);
  insertDefaultAccounts(ss);
  insertDefaultAdmin(ss);

  // 初期シート（Sheet1等）を削除
  try {
    const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
  } catch (e) { /* 無視 */ }

  ui.alert('初期化完了', `作成: ${created}シート / スキップ: ${skipped}シート（既存）\n\nデフォルト管理者:\nID: admin / PW: admin123\n\n※ ログイン後、パスワードを必ず変更してください。`, ui.ButtonSet.OK);
}

// === ヘッダー行フォーマット ===
function formatHeaderRow(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#4a90d9')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 列幅自動調整
  for (let i = 1; i <= numCols; i++) {
    sheet.setColumnWidth(i, 120);
  }
}

// === デフォルト設定投入 ===
function insertDefaultSettings(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet || sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
}

// === デフォルト勤怠設定投入 ===
function insertDefaultAttendanceSettings(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE_SETTINGS);
  if (!sheet || sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, DEFAULT_ATTENDANCE_SETTINGS.length, 3).setValues(DEFAULT_ATTENDANCE_SETTINGS);
}

// === デフォルト勘定科目投入 ===
function insertDefaultAccounts(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNTS);
  if (!sheet || sheet.getLastRow() > 1) return;
  const now = formatDate(getNow());
  const rows = DEFAULT_ACCOUNTS.map(a => [...a, now, now]);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

// === デフォルト管理者作成 ===
function insertDefaultAdmin(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.AUTH);
  if (!sheet || sheet.getLastRow() > 1) return;
  const now = formatDateTime(getNow());
  sheet.appendRow([
    'admin',
    hashPassword('admin123'),
    '管理者',
    'admin',
    '',
    true,
    '',
    now,
    now
  ]);
}

// === シートリセット（開発用・本番では使わない） ===
function resetAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ 全シート削除',
    'すべてのシートとデータを削除します。\nこの操作は取り消せません。\n本当に実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 最低1シート残す必要があるので仮シート作成
  const tempSheet = ss.insertSheet('_temp_');
  ss.getSheets().forEach(s => {
    if (s.getName() !== '_temp_') {
      try { ss.deleteSheet(s); } catch (e) {}
    }
  });
  tempSheet.setName('Sheet1');
  ui.alert('全シートを削除しました。initializeSystem() で再初期化してください。');
}