/**
 * 経理入力システム v2.0 - 認証・セッション管理
 */

// === ログイン ===
function login(userId, password) {
  try {
    const result = findRowByValue(SHEET_NAMES.AUTH, 0, userId);
    if (!result) return { success: false, error: 'ユーザーIDまたはパスワードが正しくありません' };

    const row = result.data;
    const storedHash = row[1];
    const userName = row[2];
    const role = row[3];
    const active = row[5];

    if (active === false || active === 'FALSE' || active === 'false') {
      return { success: false, error: 'このアカウントは無効です' };
    }

    if (hashPassword(password) !== storedHash) {
      return { success: false, error: 'ユーザーIDまたはパスワードが正しくありません' };
    }

    // セッション作成
    const token = generateSessionToken();
    const sessionData = JSON.stringify({
      userId: userId,
      userName: userName,
      role: role,
      loginTime: formatDateTime(getNow())
    });
    CacheService.getScriptCache().put(SESSION_PREFIX + token, sessionData, SESSION_DURATION);

    // 最終ログイン更新
    const sheet = getSheet(SHEET_NAMES.AUTH);
    sheet.getRange(result.row, 7).setValue(formatDateTime(getNow()));

    writeOperationLog(userId, userName, 'ログイン', '認証', 'ログイン成功');

    return {
      success: true,
      token: token,
      user: { userId, userName, role }
    };
  } catch (e) {
    console.error('ログインエラー:', e);
    return { success: false, error: 'システムエラーが発生しました' };
  }
}

// === ログアウト ===
function logout(token) {
  try {
    const session = validateSession(token);
    if (session.valid) {
      writeOperationLog(session.user.userId, session.user.userName, 'ログアウト', '認証', '');
      CacheService.getScriptCache().remove(SESSION_PREFIX + token);
    }
    return { success: true };
  } catch (e) {
    return { success: true }; // ログアウトは常に成功扱い
  }
}

// === セッション検証 ===
function validateSession(token) {
  if (!token) return { valid: false };
  try {
    const cached = CacheService.getScriptCache().get(SESSION_PREFIX + token);
    if (!cached) return { valid: false };
    const sessionData = JSON.parse(cached);
    // セッション延長
    CacheService.getScriptCache().put(SESSION_PREFIX + token, cached, SESSION_DURATION);
    return { valid: true, user: sessionData };
  } catch (e) {
    return { valid: false };
  }
}

// === パスワード変更 ===
function changePassword(token, oldPassword, newPassword) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  if (!newPassword || newPassword.length < 6) {
    return { success: false, error: 'パスワードは6文字以上で入力してください' };
  }

  const result = findRowByValue(SHEET_NAMES.AUTH, 0, session.user.userId);
  if (!result) return { success: false, error: 'ユーザーが見つかりません' };

  if (hashPassword(oldPassword) !== result.data[1]) {
    return { success: false, error: '現在のパスワードが正しくありません' };
  }

  const sheet = getSheet(SHEET_NAMES.AUTH);
  sheet.getRange(result.row, 2).setValue(hashPassword(newPassword));
  sheet.getRange(result.row, 9).setValue(formatDateTime(getNow()));

  writeOperationLog(session.user.userId, session.user.userName, 'パスワード変更', '認証', '');
  return { success: true };
}

// === ユーザー管理（管理者用） ===
function getUsers(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '権限がありません' };

  const data = getSheetData(SHEET_NAMES.AUTH);
  // パスワードハッシュを除外
  const users = data.map(u => ({
    userId: u['ユーザーID'],
    userName: u['ユーザー名'],
    role: u['権限'],
    email: u['メールアドレス'],
    active: u['有効フラグ'],
    lastLogin: u['最終ログイン'],
    createdAt: u['作成日']
  }));
  return { success: true, data: users };
}

function addUser(token, userData) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '権限がありません' };

  if (!userData.userId || !userData.password || !userData.userName) {
    return { success: false, error: 'ユーザーID、パスワード、ユーザー名は必須です' };
  }

  // 重複チェック
  const existing = findRowByValue(SHEET_NAMES.AUTH, 0, userData.userId);
  if (existing) return { success: false, error: 'このユーザーIDは既に使用されています' };

  const now = formatDateTime(getNow());
  appendRow(SHEET_NAMES.AUTH, [
    userData.userId,
    hashPassword(userData.password),
    userData.userName,
    userData.role || ROLES.EDITOR,
    userData.email || '',
    true,
    '',
    now,
    now
  ]);

  writeOperationLog(session.user.userId, session.user.userName, 'ユーザー追加', '認証',
    `ユーザー追加: ${userData.userId} (${userData.userName})`);
  return { success: true };
}

function updateUser(token, userId, userData) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '権限がありません' };

  const result = findRowByValue(SHEET_NAMES.AUTH, 0, userId);
  if (!result) return { success: false, error: 'ユーザーが見つかりません' };

  const sheet = getSheet(SHEET_NAMES.AUTH);
  const row = result.row;

  if (userData.userName) sheet.getRange(row, 3).setValue(userData.userName);
  if (userData.role) sheet.getRange(row, 4).setValue(userData.role);
  if (userData.email !== undefined) sheet.getRange(row, 5).setValue(userData.email);
  if (userData.active !== undefined) sheet.getRange(row, 6).setValue(userData.active);
  if (userData.password) sheet.getRange(row, 2).setValue(hashPassword(userData.password));
  sheet.getRange(row, 9).setValue(formatDateTime(getNow()));

  writeOperationLog(session.user.userId, session.user.userName, 'ユーザー更新', '認証',
    `ユーザー更新: ${userId}`);
  return { success: true };
}

function deleteUser(token, userId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '権限がありません' };
  if (userId === session.user.userId) return { success: false, error: '自分自身は削除できません' };

  const result = findRowByValue(SHEET_NAMES.AUTH, 0, userId);
  if (!result) return { success: false, error: 'ユーザーが見つかりません' };

  // 論理削除
  const sheet = getSheet(SHEET_NAMES.AUTH);
  sheet.getRange(result.row, 6).setValue(false);
  sheet.getRange(result.row, 9).setValue(formatDateTime(getNow()));

  writeOperationLog(session.user.userId, session.user.userName, 'ユーザー無効化', '認証',
    `ユーザー無効化: ${userId}`);
  return { success: true };
}