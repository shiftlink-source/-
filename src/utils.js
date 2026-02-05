/**
 * 経理入力システム v2.0 - ユーティリティ
 */

// === ID生成 ===
function generateId(prefix) {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return prefix ? `${prefix}-${timestamp}-${random}` : `${timestamp}-${random}`;
}

function generateSequentialId(sheet, column, prefix) {
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][column]);
    if (id.startsWith(prefix)) {
      const num = parseInt(id.replace(prefix, ''), 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return prefix + String(maxNum + 1).padStart(3, '0');
}

// === 日付フォーマット ===
function formatDate(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function formatDateTime(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

function formatTime(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'HH:mm');
}

function getNow() {
  return new Date();
}

function getTodayString() {
  return formatDate(getNow());
}

// === パスワードハッシュ ===
function hashPassword(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return raw.map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');
}

// === セッショントークン生成 ===
function generateSessionToken() {
  const bytes = Utilities.getUuid();
  return bytes;
}

// === シート取得ヘルパー ===
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。initializeSystem()を実行してください。`);
  return sheet;
}

// === シートデータ取得（ヘッダー付きオブジェクト配列） ===
function getSheetData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map((row, idx) => {
    const obj = { _row: idx + 2 }; // 実際の行番号（1始まり+ヘッダー分）
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// === シートに行追加 ===
function appendRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  sheet.appendRow(rowData);
  return sheet.getLastRow();
}

// === シートの特定行を更新 ===
function updateRow(sheetName, rowNumber, rowData) {
  const sheet = getSheet(sheetName);
  const numCols = rowData.length;
  sheet.getRange(rowNumber, 1, 1, numCols).setValues([rowData]);
}

// === 検索（単一列） ===
function findRowByValue(sheetName, columnIndex, value) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][columnIndex]) === String(value)) {
      return { row: i + 1, data: data[i], headers: data[0] };
    }
  }
  return null;
}

// === 数値フォーマット ===
function formatCurrency(num) {
  if (num === null || num === undefined || num === '') return '¥0';
  return '¥' + Number(num).toLocaleString('ja-JP');
}

// === バリデーション ===
function isValidDate(str) {
  if (!str) return false;
  const d = new Date(str);
  return !isNaN(d.getTime());
}

function isValidTime(str) {
  if (!str) return false;
  return /^\d{1,2}:\d{2}$/.test(str);
}

function isNotEmpty(value) {
  return value !== null && value !== undefined && String(value).trim() !== '';
}

// === オブジェクトをシート行データに変換 ===
function objectToRow(obj, headers) {
  return headers.map(h => obj[h] !== undefined ? obj[h] : '');
}

// === シートのヘッダー取得 ===
function getSheetHeaders(sheetName) {
  const sheet = getSheet(sheetName);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}