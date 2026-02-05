/**
 * 経理入力システム v2.0 - マスタ管理サービス
 */

// ============================================================
// 勘定科目マスタ
// ============================================================

function getAccounts(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.ACCOUNTS);
  const accounts = data.map(r => ({
    code: String(r['科目コード']),
    name: r['科目名'],
    category: r['カテゴリ'],
    subAccount: r['補助科目'],
    taxType: r['税区分'],
    order: r['表示順'],
    active: r['有効フラグ']
  }));
  return { success: true, data: accounts };
}

function addAccount(token, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  if (!data.code || !data.name || !data.category) {
    return { success: false, error: '科目コード、科目名、カテゴリは必須です' };
  }

  const existing = findRowByValue(SHEET_NAMES.ACCOUNTS, 0, data.code);
  if (existing) return { success: false, error: 'この科目コードは既に使用されています' };

  const now = formatDate(getNow());
  appendRow(SHEET_NAMES.ACCOUNTS, [
    data.code, data.name, data.category, data.subAccount || '',
    data.taxType || '対象外', data.order || 9999, true, now, now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '追加', '勘定科目',
    `${data.code}: ${data.name}`);
  return { success: true };
}

function updateAccount(token, code, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.ACCOUNTS, 0, code);
  if (!result) return { success: false, error: '勘定科目が見つかりません' };

  const now = formatDate(getNow());
  updateRow(SHEET_NAMES.ACCOUNTS, result.row, [
    code,
    data.name || result.data[1],
    data.category || result.data[2],
    data.subAccount !== undefined ? data.subAccount : result.data[3],
    data.taxType || result.data[4],
    data.order !== undefined ? data.order : result.data[5],
    data.active !== undefined ? data.active : result.data[6],
    result.data[7], // 登録日は変更しない
    now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '更新', '勘定科目',
    `${code}: ${data.name || result.data[1]}`);
  return { success: true };
}

function deleteAccount(token, code) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '管理者権限が必要です' };

  const result = findRowByValue(SHEET_NAMES.ACCOUNTS, 0, code);
  if (!result) return { success: false, error: '勘定科目が見つかりません' };

  // 論理削除
  const sheet = getSheet(SHEET_NAMES.ACCOUNTS);
  sheet.getRange(result.row, 7).setValue(false);
  sheet.getRange(result.row, 9).setValue(formatDate(getNow()));

  writeOperationLog(session.user.userId, session.user.userName, '無効化', '勘定科目', `${code}`);
  return { success: true };
}

// ============================================================
// 取引先マスタ
// ============================================================

function getPartners(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.PARTNERS);
  return { success: true, data: data.map(r => ({
    id: r['取引先ID'],
    name: r['取引先名'],
    nameKana: r['取引先名カナ'],
    type: r['区分'],
    zip: r['郵便番号'],
    address: r['住所'],
    tel: r['電話番号'],
    email: r['メール'],
    bankName: r['振込先銀行'],
    bankBranch: r['振込先支店'],
    accountType: r['口座種別'],
    accountNumber: r['口座番号'],
    accountHolder: r['口座名義'],
    invoiceNo: r['インボイス登録番号'],
    closingDay: r['締め日'],
    paymentTerms: r['支払条件'],
    active: r['有効フラグ']
  }))};
}

function addPartner(token, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  if (!data.name || !data.type) {
    return { success: false, error: '取引先名と区分は必須です' };
  }

  const sheet = getSheet(SHEET_NAMES.PARTNERS);
  const id = generateSequentialId(sheet, 0, 'P');
  const now = formatDate(getNow());

  appendRow(SHEET_NAMES.PARTNERS, [
    id, data.name, data.nameKana || '', data.type,
    data.zip || '', data.address || '', data.tel || '', data.email || '',
    data.bankName || '', data.bankBranch || '', data.accountType || '',
    data.accountNumber || '', data.accountHolder || '', data.invoiceNo || '',
    data.closingDay || '', data.paymentTerms || '', true, now, now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '追加', '取引先',
    `${id}: ${data.name}`);
  return { success: true, id: id };
}

function updatePartner(token, id, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.PARTNERS, 0, id);
  if (!result) return { success: false, error: '取引先が見つかりません' };

  const old = result.data;
  const now = formatDate(getNow());
  updateRow(SHEET_NAMES.PARTNERS, result.row, [
    id,
    data.name || old[1], data.nameKana !== undefined ? data.nameKana : old[2],
    data.type || old[3], data.zip !== undefined ? data.zip : old[4],
    data.address !== undefined ? data.address : old[5],
    data.tel !== undefined ? data.tel : old[6], data.email !== undefined ? data.email : old[7],
    data.bankName !== undefined ? data.bankName : old[8],
    data.bankBranch !== undefined ? data.bankBranch : old[9],
    data.accountType !== undefined ? data.accountType : old[10],
    data.accountNumber !== undefined ? data.accountNumber : old[11],
    data.accountHolder !== undefined ? data.accountHolder : old[12],
    data.invoiceNo !== undefined ? data.invoiceNo : old[13],
    data.closingDay !== undefined ? data.closingDay : old[14],
    data.paymentTerms !== undefined ? data.paymentTerms : old[15],
    data.active !== undefined ? data.active : old[16],
    old[17], // 登録日
    now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '更新', '取引先',
    `${id}: ${data.name || old[1]}`);
  return { success: true };
}

function deletePartner(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '管理者権限が必要です' };

  const result = findRowByValue(SHEET_NAMES.PARTNERS, 0, id);
  if (!result) return { success: false, error: '取引先が見つかりません' };

  const sheet = getSheet(SHEET_NAMES.PARTNERS);
  sheet.getRange(result.row, 17).setValue(false);
  sheet.getRange(result.row, 19).setValue(formatDate(getNow()));

  writeOperationLog(session.user.userId, session.user.userName, '無効化', '取引先', id);
  return { success: true };
}

// ============================================================
// 口座マスタ
// ============================================================

function getBankAccounts(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.BANK_ACCOUNTS);
  return { success: true, data: data.map(r => ({
    id: r['口座ID'],
    bankName: r['金融機関名'],
    branchName: r['支店名'],
    accountType: r['口座種別'],
    accountNumber: r['口座番号'],
    accountHolder: r['口座名義'],
    csvFormat: r['CSVフォーマット'],
    usage: r['用途'],
    active: r['有効フラグ']
  }))};
}

function addBankAccount(token, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  if (!data.bankName) return { success: false, error: '金融機関名は必須です' };

  const sheet = getSheet(SHEET_NAMES.BANK_ACCOUNTS);
  const id = generateSequentialId(sheet, 0, 'BA');
  const now = formatDate(getNow());

  appendRow(SHEET_NAMES.BANK_ACCOUNTS, [
    id, data.bankName, data.branchName || '', data.accountType || '普通',
    data.accountNumber || '', data.accountHolder || '', data.csvFormat || '',
    data.usage || '', true, now, now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '追加', '口座',
    `${id}: ${data.bankName}`);
  return { success: true, id: id };
}

function updateBankAccount(token, id, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.BANK_ACCOUNTS, 0, id);
  if (!result) return { success: false, error: '口座が見つかりません' };

  const old = result.data;
  const now = formatDate(getNow());
  updateRow(SHEET_NAMES.BANK_ACCOUNTS, result.row, [
    id,
    data.bankName || old[1], data.branchName !== undefined ? data.branchName : old[2],
    data.accountType || old[3], data.accountNumber !== undefined ? data.accountNumber : old[4],
    data.accountHolder !== undefined ? data.accountHolder : old[5],
    data.csvFormat !== undefined ? data.csvFormat : old[6],
    data.usage !== undefined ? data.usage : old[7],
    data.active !== undefined ? data.active : old[8],
    old[9], now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '更新', '口座',
    `${id}: ${data.bankName || old[1]}`);
  return { success: true };
}

function deleteBankAccount(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '管理者権限が必要です' };

  const result = findRowByValue(SHEET_NAMES.BANK_ACCOUNTS, 0, id);
  if (!result) return { success: false, error: '口座が見つかりません' };

  const sheet = getSheet(SHEET_NAMES.BANK_ACCOUNTS);
  sheet.getRange(result.row, 9).setValue(false);
  sheet.getRange(result.row, 11).setValue(formatDate(getNow()));

  writeOperationLog(session.user.userId, session.user.userName, '無効化', '口座', id);
  return { success: true };
}

// ============================================================
// 設定
// ============================================================

function getSettings(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.SETTINGS);
  const settings = {};
  data.forEach(r => { settings[r['設定キー']] = r['設定値']; });
  return { success: true, data: settings };
}

function updateSettings(token, settings) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '管理者権限が必要です' };

  const sheet = getSheet(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();

  Object.keys(settings).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(settings[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, settings[key], '']);
    }
  });

  writeOperationLog(session.user.userId, session.user.userName, '更新', '設定',
    `更新キー: ${Object.keys(settings).join(', ')}`);
  return { success: true };
}

// ============================================================
// ダッシュボード
// ============================================================

function getDashboardData(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  try {
    const accounts = getSheetData(SHEET_NAMES.ACCOUNTS);
    const partners = getSheetData(SHEET_NAMES.PARTNERS);
    const bankAccounts = getSheetData(SHEET_NAMES.BANK_ACCOUNTS);
    const pettyCash = getSheetData(SHEET_NAMES.PETTY_CASH);

    // 小口現金の今月データ
    const now = new Date();
    const thisMonth = `${now.getFullYear()}/${String(now.getMonth() + 1).padStart(2, '0')}`;
    const monthlyPettyCash = pettyCash.filter(r => {
      const d = String(r['日付']);
      return d.startsWith(thisMonth);
    });

    const totalIn = monthlyPettyCash.reduce((sum, r) => sum + (Number(r['入金額']) || 0), 0);
    const totalOut = monthlyPettyCash.reduce((sum, r) => sum + (Number(r['出金額']) || 0), 0);

    // 小口現金の最新残高
    const latestBalance = pettyCash.length > 0 ? Number(pettyCash[pettyCash.length - 1]['残高']) || 0 : 0;

    return {
      success: true,
      data: {
        accountCount: accounts.filter(a => a['有効フラグ'] === true || a['有効フラグ'] === 'TRUE').length,
        partnerCount: partners.filter(p => p['有効フラグ'] === true || p['有効フラグ'] === 'TRUE').length,
        bankAccountCount: bankAccounts.filter(b => b['有効フラグ'] === true || b['有効フラグ'] === 'TRUE').length,
        pettyCashBalance: latestBalance,
        monthlyPettyCashIn: totalIn,
        monthlyPettyCashOut: totalOut,
        monthlyPettyCashCount: monthlyPettyCash.length,
        recentPettyCash: monthlyPettyCash.slice(-5).reverse().map(r => ({
          date: r['日付'],
          summary: r['摘要'],
          inAmount: r['入金額'],
          outAmount: r['出金額'],
          balance: r['残高']
        }))
      }
    };
  } catch (e) {
    console.error('ダッシュボードエラー:', e);
    return { success: true, data: {
      accountCount: 0, partnerCount: 0, bankAccountCount: 0,
      pettyCashBalance: 0, monthlyPettyCashIn: 0, monthlyPettyCashOut: 0,
      monthlyPettyCashCount: 0, recentPettyCash: []
    }};
  }
}