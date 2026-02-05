/**
 * 経理入力システム v2.0 - 小口現金管理
 */

function getPettyCashEntries(token, filters) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.PETTY_CASH);
  let results = data.map(r => ({
    id: r['取引ID'],
    date: r['日付'],
    account: r['勘定科目'],
    subAccount: r['補助科目'],
    partner: r['取引先'],
    summary: r['摘要'],
    inAmount: r['入金額'],
    outAmount: r['出金額'],
    balance: r['残高'],
    taxType: r['税区分'],
    taxAmount: r['税額'],
    registeredBy: r['登録者'],
    registeredAt: r['登録日時'],
    note: r['備考'],
    _row: r._row
  }));

  if (filters) {
    if (filters.month) {
      results = results.filter(r => String(r.date).startsWith(filters.month));
    }
    if (filters.account) {
      results = results.filter(r => r.account === filters.account);
    }
    if (filters.partner) {
      results = results.filter(r => r.partner === filters.partner);
    }
  }

  return { success: true, data: results };
}

function addPettyCashEntry(token, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  // バリデーション
  if (!data.date) return { success: false, error: '日付は必須です' };
  if (!data.account) return { success: false, error: '勘定科目は必須です' };
  if (!data.summary) return { success: false, error: '摘要は必須です' };
  if ((!data.inAmount && !data.outAmount) || (data.inAmount && data.outAmount)) {
    return { success: false, error: '入金額または出金額のどちらか一方を入力してください' };
  }

  const inAmt = Number(data.inAmount) || 0;
  const outAmt = Number(data.outAmount) || 0;

  // 残高計算（直前の残高を取得）
  const existingData = getSheetData(SHEET_NAMES.PETTY_CASH);
  const prevBalance = existingData.length > 0
    ? Number(existingData[existingData.length - 1]['残高']) || 0
    : 0;
  const newBalance = prevBalance + inAmt - outAmt;

  // 税額計算
  const taxAmount = calculateTax(inAmt || outAmt, data.taxType);

  const sheet = getSheet(SHEET_NAMES.PETTY_CASH);
  const id = generateSequentialId(sheet, 0, 'PC');
  const now = formatDateTime(getNow());

  appendRow(SHEET_NAMES.PETTY_CASH, [
    id, data.date, data.account, data.subAccount || '', data.partner || '',
    data.summary, inAmt || '', outAmt || '', newBalance,
    data.taxType || '対象外', taxAmount, session.user.userName, now, data.note || ''
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '追加', '小口現金',
    `${id}: ${data.summary} ${inAmt ? '入金' + inAmt : '出金' + outAmt}`);
  return { success: true, id: id, balance: newBalance };
}

function updatePettyCashEntry(token, id, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.PETTY_CASH, 0, id);
  if (!result) return { success: false, error: '取引が見つかりません' };

  const old = result.data;
  const inAmt = data.inAmount !== undefined ? (Number(data.inAmount) || 0) : (Number(old[6]) || 0);
  const outAmt = data.outAmount !== undefined ? (Number(data.outAmount) || 0) : (Number(old[7]) || 0);
  const taxType = data.taxType || old[9];
  const taxAmount = calculateTax(inAmt || outAmt, taxType);

  updateRow(SHEET_NAMES.PETTY_CASH, result.row, [
    id,
    data.date || old[1],
    data.account || old[2],
    data.subAccount !== undefined ? data.subAccount : old[3],
    data.partner !== undefined ? data.partner : old[4],
    data.summary || old[5],
    inAmt || '', outAmt || '',
    old[8], // 残高は再計算が必要（下記で対応）
    taxType, taxAmount,
    old[11], old[12], // 登録者・登録日時は変更しない
    data.note !== undefined ? data.note : old[13]
  ]);

  // 当該行以降の残高を再計算
  recalculateBalances(result.row);

  writeOperationLog(session.user.userId, session.user.userName, '更新', '小口現金',
    `${id}: ${data.summary || old[5]}`);
  return { success: true };
}

function deletePettyCashEntry(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.PETTY_CASH, 0, id);
  if (!result) return { success: false, error: '取引が見つかりません' };

  const sheet = getSheet(SHEET_NAMES.PETTY_CASH);
  const deletedRow = result.row;
  sheet.deleteRow(deletedRow);

  // 残高を再計算
  recalculateBalances(deletedRow);

  writeOperationLog(session.user.userId, session.user.userName, '削除', '小口現金',
    `${id}: ${result.data[5]}`);
  return { success: true };
}

// === 残高再計算 ===
function recalculateBalances(fromRow) {
  const sheet = getSheet(SHEET_NAMES.PETTY_CASH);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
  let balance = 0;

  for (let i = 0; i < data.length; i++) {
    const inAmt = Number(data[i][6]) || 0;
    const outAmt = Number(data[i][7]) || 0;
    balance = balance + inAmt - outAmt;
    data[i][8] = balance;
  }

  // 残高列を一括更新
  const balances = data.map(r => [r[8]]);
  sheet.getRange(2, 9, balances.length, 1).setValues(balances);
}

// === 消費税計算 ===
function calculateTax(amount, taxType) {
  if (!amount || !taxType) return 0;
  let rate = 0;
  if (taxType.includes('10%')) rate = 0.1;
  else if (taxType.includes('8%') || taxType.includes('軽減')) rate = 0.08;
  else return 0;

  // 税込金額から税額を算出（内税計算）
  return Math.floor(amount * rate / (1 + rate));
}

// === 勘定科目リスト取得（プルダウン用） ===
function getAccountList(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.ACCOUNTS);
  return {
    success: true,
    data: data
      .filter(r => r['有効フラグ'] === true || r['有効フラグ'] === 'TRUE')
      .map(r => ({
        code: String(r['科目コード']),
        name: r['科目名'],
        category: r['カテゴリ'],
        taxType: r['税区分']
      }))
      .sort((a, b) => a.code.localeCompare(b.code))
  };
}

// === 取引先リスト取得（プルダウン用） ===
function getPartnerList(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.PARTNERS);
  return {
    success: true,
    data: data
      .filter(r => r['有効フラグ'] === true || r['有効フラグ'] === 'TRUE')
      .map(r => ({ id: r['取引先ID'], name: r['取引先名'], type: r['区分'] }))
  };
}