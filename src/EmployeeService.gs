/**
 * 経理入力システム v2.0 - 従業員マスタ管理
 */

// === 従業員一覧取得 ===
function getEmployees(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.EMPLOYEES);
  return {
    success: true,
    data: data.map(r => ({
      id: r['従業員ID'],
      name: r['従業員名'],
      nameKana: r['従業員名カナ'],
      department: r['所属'],
      employmentType: r['雇用形態'],
      hourlyRate: r['時給'],
      monthlyRate: r['月給'],
      lineUserId: r['LINE UserID'],
      wechatOpenId: r['WeChat OpenID'],
      whatsappPhone: r['WhatsApp電話番号'],
      mainApp: r['メインアプリ'],
      registrationCode: r['登録コード'],
      active: r['有効フラグ'],
      createdAt: r['登録日'],
      _row: r._row
    }))
  };
}

// === 従業員追加 ===
function addEmployee(token, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  if (!data.name) return { success: false, error: '従業員名は必須です' };
  if (!data.employmentType) return { success: false, error: '雇用形態は必須です' };

  const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
  const id = generateSequentialId(sheet, 0, 'E');
  const regCode = generateRegistrationCode();
  const now = formatDate(getNow());

  appendRow(SHEET_NAMES.EMPLOYEES, [
    id, data.name, data.nameKana || '', data.department || '',
    data.employmentType, data.hourlyRate || '', data.monthlyRate || '',
    data.lineUserId || '', data.wechatOpenId || '', data.whatsappPhone || '',
    data.mainApp || '', regCode, true, now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '追加', '従業員',
    `${id}: ${data.name}`);
  return { success: true, id: id, registrationCode: regCode };
}

// === 従業員更新 ===
function updateEmployee(token, id, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.EMPLOYEES, 0, id);
  if (!result) return { success: false, error: '従業員が見つかりません' };

  const old = result.data;
  updateRow(SHEET_NAMES.EMPLOYEES, result.row, [
    id,
    data.name || old[1],
    data.nameKana !== undefined ? data.nameKana : old[2],
    data.department !== undefined ? data.department : old[3],
    data.employmentType || old[4],
    data.hourlyRate !== undefined ? data.hourlyRate : old[5],
    data.monthlyRate !== undefined ? data.monthlyRate : old[6],
    data.lineUserId !== undefined ? data.lineUserId : old[7],
    data.wechatOpenId !== undefined ? data.wechatOpenId : old[8],
    data.whatsappPhone !== undefined ? data.whatsappPhone : old[9],
    data.mainApp !== undefined ? data.mainApp : old[10],
    old[11], // 登録コードは変更しない
    data.active !== undefined ? data.active : old[12],
    old[13]  // 登録日は変更しない
  ]);

  writeOperationLog(session.user.userId, session.user.userName, '更新', '従業員',
    `${id}: ${data.name || old[1]}`);
  return { success: true };
}

// === 従業員無効化 ===
function deleteEmployee(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '管理者権限が必要です' };

  const result = findRowByValue(SHEET_NAMES.EMPLOYEES, 0, id);
  if (!result) return { success: false, error: '従業員が見つかりません' };

  const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
  sheet.getRange(result.row, 13).setValue(false); // 有効フラグ

  writeOperationLog(session.user.userId, session.user.userName, '無効化', '従業員', id);
  return { success: true };
}

// === 登録コード生成（6桁） ===
function generateRegistrationCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 紛らわしい文字を除外
  let code = '';
  for (let i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

// === 登録コード再発行 ===
function regenerateRegistrationCode(token, employeeId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.EMPLOYEES, 0, employeeId);
  if (!result) return { success: false, error: '従業員が見つかりません' };

  const newCode = generateRegistrationCode();
  const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
  sheet.getRange(result.row, 12).setValue(newCode); // 登録コード列

  writeOperationLog(session.user.userId, session.user.userName, '登録コード再発行', '従業員',
    `${employeeId}: 新コード=${newCode}`);
  return { success: true, registrationCode: newCode };
}

// === 登録コードで従業員検索（Webhook用・内部関数） ===
function findEmployeeByRegistrationCode(code) {
  const data = getSheetData(SHEET_NAMES.EMPLOYEES);
  return data.find(r => r['登録コード'] === code && (r['有効フラグ'] === true || r['有効フラグ'] === 'TRUE'));
}

// === メッセージアプリIDで従業員検索（Webhook用・内部関数） ===
function findEmployeeByAppId(source, appUserId) {
  const data = getSheetData(SHEET_NAMES.EMPLOYEES);
  let columnName;
  switch (source) {
    case 'line': columnName = 'LINE UserID'; break;
    case 'wechat': columnName = 'WeChat OpenID'; break;
    case 'whatsapp': columnName = 'WhatsApp電話番号'; break;
    default: return null;
  }
  return data.find(r =>
    String(r[columnName]) === String(appUserId) &&
    (r['有効フラグ'] === true || r['有効フラグ'] === 'TRUE')
  );
}

// === メッセージアプリIDを従業員に紐付け（内部関数） ===
function linkAppIdToEmployee(employeeRow, source, appUserId) {
  const sheet = getSheet(SHEET_NAMES.EMPLOYEES);
  let col;
  switch (source) {
    case 'line': col = 8; break;     // LINE UserID
    case 'wechat': col = 9; break;   // WeChat OpenID
    case 'whatsapp': col = 10; break; // WhatsApp電話番号
    default: return;
  }
  sheet.getRange(employeeRow, col).setValue(appUserId);
  // メインアプリが未設定なら設定
  const mainApp = sheet.getRange(employeeRow, 11).getValue();
  if (!mainApp) {
    const appName = source === 'line' ? 'LINE' : source === 'wechat' ? 'WeChat' : 'WhatsApp';
    sheet.getRange(employeeRow, 11).setValue(appName);
  }
}

// === 従業員リスト取得（プルダウン用） ===
function getEmployeeList(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.EMPLOYEES);
  return {
    success: true,
    data: data
      .filter(r => r['有効フラグ'] === true || r['有効フラグ'] === 'TRUE')
      .map(r => ({
        id: r['従業員ID'],
        name: r['従業員名'],
        department: r['所属'],
        employmentType: r['雇用形態']
      }))
  };
}
