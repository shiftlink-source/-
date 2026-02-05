/**
 * 経理入力システム v2.0 - 出退勤サービス
 */

// ============================================================
// 打刻処理（Webhook / 手動共通）
// ============================================================

/**
 * 打刻を記録する（メインエントリーポイント）
 * @param {string} employeeId - 従業員ID
 * @param {string} punchType - 出勤/退勤/休憩開始/休憩終了
 * @param {string} source - LINE/WeChat/WhatsApp/手動/CSV
 * @param {object} meta - { appUserId, rawMessage, lat, lng }
 * @returns {object} { success, message, data }
 */
function recordPunch(employeeId, punchType, source, meta) {
  try {
    const now = getNow();
    const today = formatDate(now);
    const timeStr = formatTime(now);

    // 打刻ログに記録
    const logSheet = getSheet(SHEET_NAMES.PUNCH_LOG);
    const logId = 'PL-' + now.getTime();
    logSheet.appendRow([
      logId, employeeId, punchType, formatDateTime(now), source,
      (meta && meta.appUserId) || '', (meta && meta.rawMessage) || '',
      (meta && meta.lat) || '', (meta && meta.lng) || ''
    ]);

    // 出退勤データを更新
    const result = updateAttendanceRecord(employeeId, today, punchType, timeStr, source);

    return {
      success: true,
      message: result.message,
      data: {
        punchType: punchType,
        time: timeStr,
        date: today,
        workSummary: result.workSummary || null
      }
    };
  } catch (e) {
    console.error('打刻エラー:', e);
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  }
}

/**
 * 出退勤データの更新・作成
 */
function updateAttendanceRecord(employeeId, date, punchType, timeStr, source) {
  const sheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // 従業員名を取得
  const empResult = findRowByValue(SHEET_NAMES.EMPLOYEES, 0, employeeId);
  const empName = empResult ? empResult.data[1] : employeeId;

  // 当日の既存レコードを検索
  let existingRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(employeeId) && String(data[i][3]) === String(date)) {
      existingRow = i + 1; // 1始まり
      break;
    }
  }

  let message = '';
  let workSummary = null;

  if (existingRow === -1) {
    // 新規レコード作成
    const recordId = 'AT-' + getNow().getTime();
    const newRow = new Array(headers.length).fill('');
    newRow[0] = recordId;        // レコードID
    newRow[1] = employeeId;      // 従業員ID
    newRow[2] = empName;         // 従業員名
    newRow[3] = date;            // 日付
    newRow[16] = '正常';         // ステータス
    newRow[17] = source;         // 打刻ソース
    newRow[18] = false;          // 修正フラグ

    switch (punchType) {
      case '出勤':
        newRow[4] = timeStr;
        message = `✅ 出勤を記録しました（${timeStr}）`;
        break;
      case '退勤':
        newRow[5] = timeStr;
        message = `✅ 退勤を記録しました（${timeStr}）`;
        break;
      case '休憩開始':
        newRow[6] = timeStr;
        message = `☕ 休憩開始を記録しました（${timeStr}）`;
        break;
      case '休憩終了':
        newRow[7] = timeStr;
        message = `🔙 休憩終了を記録しました（${timeStr}）`;
        break;
    }

    sheet.appendRow(newRow);
    existingRow = sheet.getLastRow();
  } else {
    // 既存レコード更新
    const row = data[existingRow - 1];

    switch (punchType) {
      case '出勤':
        if (row[4]) {
          message = `⚠️ 本日は既に出勤打刻済みです（${row[4]}）`;
          return { message };
        }
        sheet.getRange(existingRow, 5).setValue(timeStr);
        message = `✅ 出勤を記録しました（${timeStr}）`;
        break;

      case '退勤':
        sheet.getRange(existingRow, 6).setValue(timeStr);
        message = `✅ 退勤を記録しました（${timeStr}）`;
        break;

      case '休憩開始':
        // 空いている休憩スロットに記録
        if (!row[6]) { sheet.getRange(existingRow, 7).setValue(timeStr); }
        else if (!row[8]) { sheet.getRange(existingRow, 9).setValue(timeStr); }
        else if (!row[10]) { sheet.getRange(existingRow, 11).setValue(timeStr); }
        else { message = '⚠️ 休憩枠が上限（3回）に達しています'; return { message }; }
        message = `☕ 休憩開始を記録しました（${timeStr}）`;
        break;

      case '休憩終了':
        // 最後に開始された休憩の終了を記録
        if (row[10] && !row[11]) { sheet.getRange(existingRow, 12).setValue(timeStr); }
        else if (row[8] && !row[9]) { sheet.getRange(existingRow, 10).setValue(timeStr); }
        else if (row[6] && !row[7]) { sheet.getRange(existingRow, 8).setValue(timeStr); }
        else { message = '⚠️ 開始されていない休憩の終了はできません'; return { message }; }
        message = `🔙 休憩終了を記録しました（${timeStr}）`;
        break;
    }
  }

  // 時間計算を実行
  if (punchType === '退勤') {
    workSummary = calculateAndUpdateTimes(sheet, existingRow);
    if (workSummary) {
      message += `\n本日の勤務時間: ${minutesToHM(workSummary.actualWork)}（休憩: ${minutesToHM(workSummary.totalBreak)}）`;
      if (workSummary.overtime > 0) message += `\n残業: ${minutesToHM(workSummary.overtime)}`;
    }
  }

  // 休憩終了時も時間更新
  if (punchType === '休憩終了') {
    calculateAndUpdateTimes(sheet, existingRow);
  }

  return { message, workSummary };
}

// ============================================================
// 勤務時間計算
// ============================================================

function calculateAndUpdateTimes(sheet, rowNumber) {
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const clockIn = row[4];   // 出勤
  const clockOut = row[5];  // 退勤

  if (!clockIn || !clockOut) return null;

  // 設定を取得
  const settings = getAttendanceSettings();
  const standardMinutes = Number(settings.working_hours) || 480;
  const nightStart = settings.night_start || '22:00';
  const nightEnd = settings.night_end || '05:00';
  const rounding = settings.rounding || 'なし';

  // 時刻をパース
  let inMin = timeToMinutes(String(clockIn));
  let outMin = timeToMinutes(String(clockOut));

  // 丸め処理
  inMin = applyRounding(inMin, rounding, 'in');
  outMin = applyRounding(outMin, rounding, 'out');

  // 日をまたぐ場合
  if (outMin < inMin) outMin += 1440;

  // 休憩時間計算
  let totalBreak = 0;
  for (let i = 0; i < 3; i++) {
    const bStart = row[6 + i * 2];
    const bEnd = row[7 + i * 2];
    if (bStart && bEnd) {
      let bsMin = timeToMinutes(String(bStart));
      let beMin = timeToMinutes(String(bEnd));
      if (beMin < bsMin) beMin += 1440;
      totalBreak += (beMin - bsMin);
    }
  }

  // 実働時間
  const grossWork = outMin - inMin;
  const actualWork = Math.max(0, grossWork - totalBreak);

  // 残業時間
  const overtime = Math.max(0, actualWork - standardMinutes);

  // 深夜時間
  const nightMinutes = calculateNightMinutes(inMin, outMin, totalBreak, nightStart, nightEnd);

  // シートに書き込み
  sheet.getRange(rowNumber, 13).setValue(totalBreak);      // 総休憩時間(分)
  sheet.getRange(rowNumber, 14).setValue(actualWork);       // 実働時間(分)
  sheet.getRange(rowNumber, 15).setValue(overtime);         // 残業時間(分)
  sheet.getRange(rowNumber, 16).setValue(nightMinutes);     // 深夜時間(分)

  return { totalBreak, actualWork, overtime, nightMinutes };
}

// === 深夜時間計算（22:00-05:00） ===
function calculateNightMinutes(inMin, outMin, breakMin, nightStartStr, nightEndStr) {
  const ns = timeToMinutes(nightStartStr);  // 22:00 = 1320
  let ne = timeToMinutes(nightEndStr);      // 05:00 = 300
  if (ne < ns) ne += 1440;                  // 翌日 = 1740

  let nightWork = 0;

  // 夜の時間帯（22:00〜29:00=翌5:00）と勤務時間の重なりを計算
  const overlapStart = Math.max(inMin, ns);
  const overlapEnd = Math.min(outMin, ne);
  if (overlapEnd > overlapStart) {
    nightWork = overlapEnd - overlapStart;
  }

  // 深夜時間帯の休憩分を概算で差し引く（簡易計算）
  // 厳密にやるなら休憩の開始終了と深夜帯の重なりを計算するが、ここでは比率で概算
  if (nightWork > 0 && breakMin > 0 && (outMin - inMin) > 0) {
    const nightRatio = nightWork / (outMin - inMin);
    nightWork = Math.max(0, Math.round(nightWork - breakMin * nightRatio));
  }

  return nightWork;
}

// === 打刻丸め処理 ===
function applyRounding(minutes, rounding, type) {
  if (rounding === 'なし') return minutes;
  if (rounding === '15分単位切り捨て') {
    if (type === 'in') return Math.ceil(minutes / 15) * 15;   // 出勤は切り上げ
    return Math.floor(minutes / 15) * 15;                      // 退勤は切り捨て
  }
  if (rounding === '15分単位四捨五入') {
    return Math.round(minutes / 15) * 15;
  }
  return minutes;
}

// ============================================================
// 出退勤データ取得（管理画面用）
// ============================================================

function getAttendanceData(token, filters) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.ATTENDANCE);
  let results = data.map(r => ({
    id: r['レコードID'],
    employeeId: r['従業員ID'],
    employeeName: r['従業員名'],
    date: r['日付'],
    clockIn: r['出勤時刻'],
    clockOut: r['退勤時刻'],
    break1Start: r['休憩1_開始'],
    break1End: r['休憩1_終了'],
    break2Start: r['休憩2_開始'],
    break2End: r['休憩2_終了'],
    break3Start: r['休憩3_開始'],
    break3End: r['休憩3_終了'],
    totalBreak: r['総休憩時間(分)'],
    actualWork: r['実働時間(分)'],
    overtime: r['残業時間(分)'],
    nightWork: r['深夜時間(分)'],
    status: r['ステータス'],
    source: r['打刻ソース'],
    modified: r['修正フラグ'],
    modifiedBy: r['修正者'],
    modifiedAt: r['修正日時'],
    modifyReason: r['修正理由'],
    note: r['備考'],
    _row: r._row
  }));

  // フィルタ適用
  if (filters) {
    if (filters.month) {
      results = results.filter(r => String(r.date).startsWith(filters.month));
    }
    if (filters.employeeId) {
      results = results.filter(r => r.employeeId === filters.employeeId);
    }
    if (filters.status) {
      results = results.filter(r => r.status === filters.status);
    }
  }

  // 日付の新しい順にソート
  results.sort((a, b) => {
    const dateCompare = String(b.date).localeCompare(String(a.date));
    if (dateCompare !== 0) return dateCompare;
    return String(a.employeeName).localeCompare(String(b.employeeName));
  });

  return { success: true, data: results };
}

// === 出退勤データ修正（管理者用） ===
function updateAttendanceRecord_admin(token, recordId, data) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  if (!data.modifyReason) return { success: false, error: '修正理由は必須です' };

  const result = findRowByValue(SHEET_NAMES.ATTENDANCE, 0, recordId);
  if (!result) return { success: false, error: 'レコードが見つかりません' };

  const sheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const row = result.row;
  const old = result.data;

  // 各フィールド更新
  if (data.clockIn !== undefined) sheet.getRange(row, 5).setValue(data.clockIn);
  if (data.clockOut !== undefined) sheet.getRange(row, 6).setValue(data.clockOut);
  if (data.break1Start !== undefined) sheet.getRange(row, 7).setValue(data.break1Start);
  if (data.break1End !== undefined) sheet.getRange(row, 8).setValue(data.break1End);
  if (data.break2Start !== undefined) sheet.getRange(row, 9).setValue(data.break2Start);
  if (data.break2End !== undefined) sheet.getRange(row, 10).setValue(data.break2End);
  if (data.break3Start !== undefined) sheet.getRange(row, 11).setValue(data.break3Start);
  if (data.break3End !== undefined) sheet.getRange(row, 12).setValue(data.break3End);
  if (data.status) sheet.getRange(row, 17).setValue(data.status);
  if (data.note !== undefined) sheet.getRange(row, 23).setValue(data.note);

  // 修正メタ情報
  sheet.getRange(row, 19).setValue(true);                           // 修正フラグ
  sheet.getRange(row, 20).setValue(session.user.userName);          // 修正者
  sheet.getRange(row, 21).setValue(formatDateTime(getNow()));       // 修正日時
  sheet.getRange(row, 22).setValue(data.modifyReason);             // 修正理由

  // 時間再計算
  calculateAndUpdateTimes(sheet, row);

  writeOperationLog(session.user.userId, session.user.userName, '出退勤修正', '出退勤',
    `${recordId}: ${data.modifyReason}`);
  return { success: true };
}

// === 手動打刻（管理画面から） ===
function manualPunch(token, employeeId, punchType, dateTime) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  const result = recordPunch(employeeId, punchType, '手動', {
    appUserId: session.user.userId,
    rawMessage: `手動打刻 by ${session.user.userName}`
  });

  writeOperationLog(session.user.userId, session.user.userName, '手動打刻', '出退勤',
    `${employeeId}: ${punchType}`);
  return result;
}

// === 現在のステータス取得（Webhook用） ===
function getEmployeeStatus(employeeId) {
  const today = formatDate(getNow());
  const data = getSheetData(SHEET_NAMES.ATTENDANCE);
  const todayRecord = data.find(r =>
    String(r['従業員ID']) === String(employeeId) && String(r['日付']) === today
  );

  if (!todayRecord) {
    return { status: '未出勤', detail: '本日の打刻はありません' };
  }

  const clockIn = todayRecord['出勤時刻'];
  const clockOut = todayRecord['退勤時刻'];

  // 最新の休憩状態を確認
  let onBreak = false;
  for (let i = 3; i >= 1; i--) {
    const bStart = todayRecord[`休憩${i}_開始`];
    const bEnd = todayRecord[`休憩${i}_終了`];
    if (bStart && !bEnd) { onBreak = true; break; }
  }

  let status, detail;
  if (clockOut) {
    const actual = todayRecord['実働時間(分)'];
    const brk = todayRecord['総休憩時間(分)'];
    status = '退勤済';
    detail = `出勤: ${clockIn} | 退勤: ${clockOut}\n実働: ${minutesToHM(actual)} | 休憩: ${minutesToHM(brk)}`;
  } else if (onBreak) {
    status = '休憩中';
    detail = `出勤: ${clockIn} | 休憩中`;
  } else if (clockIn) {
    status = '勤務中';
    detail = `出勤: ${clockIn}`;
    // 休憩情報追加
    const brk1s = todayRecord['休憩1_開始'];
    const brk1e = todayRecord['休憩1_終了'];
    if (brk1s && brk1e) detail += ` | 休憩: ${brk1s}〜${brk1e}`;
  } else {
    status = '未出勤';
    detail = '本日の出勤打刻はありません';
  }

  return { status, detail };
}

// === 月次サマリー取得 ===
function getMonthlyAttendanceSummary(token, filters) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const data = getSheetData(SHEET_NAMES.ATTENDANCE);
  const month = filters?.month || Utilities.formatDate(getNow(), 'Asia/Tokyo', 'yyyy/MM');

  const monthData = data.filter(r => String(r['日付']).startsWith(month));

  // 従業員別に集計
  const summary = {};
  monthData.forEach(r => {
    const eid = r['従業員ID'];
    if (!summary[eid]) {
      summary[eid] = {
        employeeId: eid,
        employeeName: r['従業員名'],
        days: 0,
        totalWork: 0,
        totalBreak: 0,
        totalOvertime: 0,
        totalNight: 0,
        punchMissing: 0
      };
    }
    summary[eid].days++;
    summary[eid].totalWork += Number(r['実働時間(分)']) || 0;
    summary[eid].totalBreak += Number(r['総休憩時間(分)']) || 0;
    summary[eid].totalOvertime += Number(r['残業時間(分)']) || 0;
    summary[eid].totalNight += Number(r['深夜時間(分)']) || 0;
    if (r['ステータス'] === '打刻漏れ') summary[eid].punchMissing++;
  });

  return { success: true, data: Object.values(summary), month };
}

// ============================================================
// ユーティリティ
// ============================================================

function timeToMinutes(timeStr) {
  if (!timeStr) return 0;
  const s = String(timeStr);
  const parts = s.split(':');
  if (parts.length < 2) return 0;
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

function minutesToHM(minutes) {
  if (!minutes && minutes !== 0) return '0:00';
  const m = Number(minutes);
  const h = Math.floor(m / 60);
  const min = m % 60;
  return `${h}時間${String(min).padStart(2, '0')}分`;
}

function getAttendanceSettings() {
  try {
    const data = getSheetData(SHEET_NAMES.ATTENDANCE_SETTINGS);
    const settings = {};
    data.forEach(r => { settings[r['設定キー']] = r['設定値']; });
    return settings;
  } catch (e) {
    return {
      working_hours: '480', break_time: '60', overtime_basis: '所定超過',
      night_start: '22:00', night_end: '05:00', closing_day: '月末',
      timezone: 'Asia/Tokyo', rounding: 'なし'
    };
  }
}

// === 勤怠設定取得（API） ===
function getAttendanceSettingsApi(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  return { success: true, data: getAttendanceSettings() };
}

// === 勤怠設定更新（API） ===
function updateAttendanceSettings(token, settings) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role !== ROLES.ADMIN) return { success: false, error: '管理者権限が必要です' };

  const sheet = getSheet(SHEET_NAMES.ATTENDANCE_SETTINGS);
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
    if (!found) sheet.appendRow([key, settings[key], '']);
  });

  writeOperationLog(session.user.userId, session.user.userName, '更新', '勤怠設定',
    `更新: ${Object.keys(settings).join(', ')}`);
  return { success: true };
}
