/**
 * 経理入力システム v2.0 - 出退勤CSV入出力
 */

// ============================================================
// CSV出力（エクスポート）
// ============================================================

function exportAttendanceCsv(token, filters) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const res = getAttendanceData(token, filters);
  if (!res.success) return res;

  const headers = [
    '従業員ID', '従業員名', '日付', '出勤時刻', '退勤時刻',
    '休憩1開始', '休憩1終了', '休憩2開始', '休憩2終了', '休憩3開始', '休憩3終了',
    '総休憩時間(分)', '実働時間(分)', '残業時間(分)', '深夜時間(分)', 'ステータス', '備考'
  ];

  let csv = headers.join(',') + '\n';

  res.data.forEach(r => {
    const row = [
      r.employeeId, r.employeeName, r.date, r.clockIn || '', r.clockOut || '',
      r.break1Start || '', r.break1End || '',
      r.break2Start || '', r.break2End || '',
      r.break3Start || '', r.break3End || '',
      r.totalBreak || 0, r.actualWork || 0, r.overtime || 0, r.nightWork || 0,
      r.status || '', r.note || ''
    ].map(v => {
      const s = String(v);
      return s.includes(',') || s.includes('"') || s.includes('\n')
        ? '"' + s.replace(/"/g, '""') + '"' : s;
    });
    csv += row.join(',') + '\n';
  });

  writeOperationLog(session.user.userId, session.user.userName, 'CSVエクスポート', '出退勤',
    `${res.data.length}件`);
  return { success: true, csv: csv, count: res.data.length };
}

// === CSVテンプレートダウンロード ===
function getAttendanceCsvTemplate(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const headers = [
    '従業員ID', '従業員名', '日付', '出勤時刻', '退勤時刻',
    '休憩1開始', '休憩1終了', '休憩2開始', '休憩2終了', '休憩3開始', '休憩3終了',
    '総休憩時間(分)', '実働時間(分)', '残業時間(分)', '深夜時間(分)', 'ステータス', '備考'
  ];

  const sample = [
    'E001', '山田太郎', '2025/05/01', '09:00', '18:00',
    '12:00', '13:00', '', '', '', '',
    '60', '420', '0', '0', '正常', ''
  ];

  return {
    success: true,
    csv: headers.join(',') + '\n' + sample.join(',') + '\n'
  };
}

// ============================================================
// CSV入力（インポート）
// ============================================================

function validateAttendanceCsv(token, csvText) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };

  const lines = csvText.split('\n').filter(l => l.trim());
  if (lines.length < 2) return { success: false, error: 'データ行がありません' };

  const headers = parseCsvLine(lines[0]);
  const requiredCols = ['従業員ID', '日付'];
  const missing = requiredCols.filter(c => !headers.includes(c));
  if (missing.length > 0) {
    return { success: false, error: `必須列が不足: ${missing.join(', ')}` };
  }

  // 従業員マスタ取得
  const employees = getSheetData(SHEET_NAMES.EMPLOYEES);
  const empIds = new Set(employees.map(e => String(e['従業員ID'])));

  const results = [];
  const errors = [];

  for (let i = 1; i < lines.length; i++) {
    const cols = parseCsvLine(lines[i]);
    const rowData = {};
    headers.forEach((h, idx) => { rowData[h] = cols[idx] || ''; });

    const rowErrors = [];

    // 従業員ID存在チェック
    if (!empIds.has(rowData['従業員ID'])) {
      rowErrors.push(`従業員ID「${rowData['従業員ID']}」が存在しません`);
    }

    // 日付フォーマットチェック
    if (!rowData['日付'] || !/^\d{4}\/\d{2}\/\d{2}$/.test(rowData['日付'])) {
      rowErrors.push('日付はYYYY/MM/DD形式で入力してください');
    }

    // 時刻フォーマットチェック
    const timeFields = ['出勤時刻', '退勤時刻', '休憩1開始', '休憩1終了', '休憩2開始', '休憩2終了', '休憩3開始', '休憩3終了'];
    timeFields.forEach(f => {
      if (rowData[f] && !/^\d{1,2}:\d{2}$/.test(rowData[f])) {
        rowErrors.push(`${f}はHH:mm形式で入力してください`);
      }
    });

    // 出勤 < 退勤 チェック
    if (rowData['出勤時刻'] && rowData['退勤時刻']) {
      const inMin = timeToMinutes(rowData['出勤時刻']);
      const outMin = timeToMinutes(rowData['退勤時刻']);
      // 日をまたぐ場合は許容
    }

    results.push({
      row: i + 1,
      data: rowData,
      errors: rowErrors,
      valid: rowErrors.length === 0
    });

    if (rowErrors.length > 0) {
      errors.push({ row: i + 1, errors: rowErrors });
    }
  }

  return {
    success: true,
    totalRows: results.length,
    validRows: results.filter(r => r.valid).length,
    errorRows: errors.length,
    errors: errors,
    preview: results.slice(0, 20)
  };
}

function importAttendanceCsv(token, csvText) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, error: '認証エラー' };
  if (session.user.role === ROLES.VIEWER) return { success: false, error: '権限がありません' };

  // バリデーション
  const validation = validateAttendanceCsv(token, csvText);
  if (!validation.success) return validation;
  if (validation.errorRows > 0) {
    return { success: false, error: `${validation.errorRows}行にエラーがあります。修正後に再インポートしてください。`, errors: validation.errors };
  }

  const lines = csvText.split('\n').filter(l => l.trim());
  const headers = parseCsvLine(lines[0]);
  const sheet = getSheet(SHEET_NAMES.ATTENDANCE);
  const existingData = sheet.getDataRange().getValues();

  let updated = 0;
  let added = 0;

  for (let i = 1; i < lines.length; i++) {
    const cols = parseCsvLine(lines[i]);
    const rowData = {};
    headers.forEach((h, idx) => { rowData[h] = cols[idx] || ''; });

    const empId = rowData['従業員ID'];
    const date = rowData['日付'];

    // 既存データ検索
    let existingRow = -1;
    for (let j = 1; j < existingData.length; j++) {
      if (String(existingData[j][1]) === empId && String(existingData[j][3]) === date) {
        existingRow = j + 1;
        break;
      }
    }

    // 従業員名を取得
    const empResult = findRowByValue(SHEET_NAMES.EMPLOYEES, 0, empId);
    const empName = empResult ? empResult.data[1] : rowData['従業員名'] || empId;

    if (existingRow > 0) {
      // 上書き更新
      sheet.getRange(existingRow, 5).setValue(rowData['出勤時刻'] || '');
      sheet.getRange(existingRow, 6).setValue(rowData['退勤時刻'] || '');
      sheet.getRange(existingRow, 7).setValue(rowData['休憩1開始'] || '');
      sheet.getRange(existingRow, 8).setValue(rowData['休憩1終了'] || '');
      sheet.getRange(existingRow, 9).setValue(rowData['休憩2開始'] || '');
      sheet.getRange(existingRow, 10).setValue(rowData['休憩2終了'] || '');
      sheet.getRange(existingRow, 11).setValue(rowData['休憩3開始'] || '');
      sheet.getRange(existingRow, 12).setValue(rowData['休憩3終了'] || '');
      sheet.getRange(existingRow, 17).setValue(rowData['ステータス'] || '修正済');
      sheet.getRange(existingRow, 18).setValue('CSV');
      sheet.getRange(existingRow, 19).setValue(true);
      sheet.getRange(existingRow, 20).setValue(session.user.userName);
      sheet.getRange(existingRow, 21).setValue(formatDateTime(getNow()));
      sheet.getRange(existingRow, 22).setValue('CSVインポートによる一括修正');
      sheet.getRange(existingRow, 23).setValue(rowData['備考'] || '');

      // 時間再計算
      calculateAndUpdateTimes(sheet, existingRow);
      updated++;
    } else {
      // 新規追加
      const recordId = 'AT-' + getNow().getTime() + '-' + i;
      const newRow = [
        recordId, empId, empName, date,
        rowData['出勤時刻'] || '', rowData['退勤時刻'] || '',
        rowData['休憩1開始'] || '', rowData['休憩1終了'] || '',
        rowData['休憩2開始'] || '', rowData['休憩2終了'] || '',
        rowData['休憩3開始'] || '', rowData['休憩3終了'] || '',
        '', '', '', '', // 時間計算は後で
        rowData['ステータス'] || '正常', 'CSV',
        false, '', '', '', rowData['備考'] || ''
      ];
      sheet.appendRow(newRow);
      const newRowNum = sheet.getLastRow();
      calculateAndUpdateTimes(sheet, newRowNum);
      added++;
    }
  }

  writeOperationLog(session.user.userId, session.user.userName, 'CSVインポート', '出退勤',
    `追加: ${added}件, 更新: ${updated}件`);
  return { success: true, added, updated, total: added + updated };
}

// === CSV行パーサー（ダブルクォート対応） ===
function parseCsvLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"' && line[i + 1] === '"') {
        current += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        result.push(current.trim());
        current = '';
      } else {
        current += ch;
      }
    }
  }
  result.push(current.trim());
  return result;
}
