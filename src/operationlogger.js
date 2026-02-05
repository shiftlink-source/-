/**
 * 経理入力システム v2.0 - 操作ログ
 */

function writeOperationLog(userId, userName, operation, target, detail) {
  try {
    const sheet = getSheet(SHEET_NAMES.OPERATION_LOG);
    const logId = 'LOG-' + new Date().getTime();
    sheet.appendRow([
      logId,
      formatDateTime(getNow()),
      userId || '',
      userName || '',
      operation,
      target,
      typeof detail === 'object' ? JSON.stringify(detail) : String(detail || ''),
      ''
    ]);
  } catch (e) {
    console.error('操作ログ書込エラー:', e.message);
  }
}

// === 操作ログ取得（API） ===
function getOperationLogs(token, filters) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.OPERATION_LOG);
  let results = data.reverse(); // 新しい順

  if (filters) {
    if (filters.startDate) {
      results = results.filter(r => r['日時'] >= filters.startDate);
    }
    if (filters.endDate) {
      results = results.filter(r => r['日時'] <= filters.endDate + ' 23:59:59');
    }
    if (filters.userId) {
      results = results.filter(r => r['ユーザーID'] === filters.userId);
    }
    if (filters.operation) {
      results = results.filter(r => r['操作種別'] === filters.operation);
    }
  }

  // 最大500件
  return { success: true, data: results.slice(0, 500) };
}