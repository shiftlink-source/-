/**
 * 経理入力システム v2.0 - Webhook ハンドラー
 * LINE / WhatsApp / WeChat のメッセージ受信・打刻処理・応答
 */

// ============================================================
// 打刻コマンドマッピング
// ============================================================

const PUNCH_COMMANDS = {
  '出勤': '出勤', 'しゅっきん': '出勤', 'in': '出勤', 'おはよう': '出勤',
  '退勤': '退勤', 'たいきん': '退勤', 'out': '退勤', 'おつかれ': '退勤',
  '休憩開始': '休憩開始', 'きゅうけい': '休憩開始', 'break': '休憩開始', '休憩': '休憩開始',
  '休憩終了': '休憩終了', 'もどり': '休憩終了', 'back': '休憩終了', '戻り': '休憩終了',
  '状態': 'STATUS', 'status': 'STATUS', '確認': 'STATUS',
  '履歴': 'HISTORY', 'history': 'HISTORY', '今月': 'HISTORY'
};

function parsePunchCommand(text) {
  if (!text) return null;
  const normalized = text.trim().toLowerCase();
  return PUNCH_COMMANDS[normalized] || null;
}

// ============================================================
// 共通メッセージ処理
// ============================================================

function processMessage(source, appUserId, messageText) {
  const text = String(messageText).trim();

  // 1. 登録コードチェック（6桁英数字）
  if (/^[A-Z0-9]{6}$/.test(text.toUpperCase())) {
    return handleRegistration(source, appUserId, text.toUpperCase());
  }

  // 2. 従業員を特定
  const employee = findEmployeeByAppId(source, appUserId);
  if (!employee) {
    return '⚠️ 未登録のアカウントです。\n管理者から発行された6桁の登録コードを送信してください。';
  }

  const employeeId = employee['従業員ID'];
  const employeeName = employee['従業員名'];

  // 3. コマンド解析
  const command = parsePunchCommand(text);
  if (!command) {
    return `こんにちは、${employeeName}さん。\n以下のコマンドが使えます:\n\n🟢 出勤\n🔴 退勤\n☕ 休憩開始（休憩）\n🔙 休憩終了（戻り）\n📋 状態\n📊 履歴`;
  }

  // 4. ステータス照会
  if (command === 'STATUS') {
    const status = getEmployeeStatus(employeeId);
    return `📋 ${employeeName}さんの状態\nステータス: ${status.status}\n${status.detail}`;
  }

  // 5. 履歴照会
  if (command === 'HISTORY') {
    return getMonthlyHistoryMessage(employeeId, employeeName);
  }

  // 6. 打刻実行
  const result = recordPunch(employeeId, command, source, {
    appUserId: appUserId,
    rawMessage: text
  });

  return result.message || '処理が完了しました';
}

// === 登録コード処理 ===
function handleRegistration(source, appUserId, code) {
  const employee = findEmployeeByRegistrationCode(code);
  if (!employee) {
    return '❌ 登録コードが正しくありません。\n管理者に確認してください。';
  }

  // 既に紐付け済みかチェック
  const existing = findEmployeeByAppId(source, appUserId);
  if (existing) {
    return `ℹ️ ${existing['従業員名']}さんとして既に登録されています。`;
  }

  // 紐付け
  linkAppIdToEmployee(employee._row, source, appUserId);

  const empName = employee['従業員名'];
  return `✅ 登録完了！\n${empName}さん、ようこそ。\n\n以下のコマンドで打刻できます:\n🟢 出勤\n🔴 退勤\n☕ 休憩（休憩開始）\n🔙 戻り（休憩終了）\n📋 状態`;
}

// === 月次履歴メッセージ ===
function getMonthlyHistoryMessage(employeeId, employeeName) {
  const now = getNow();
  const month = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');
  const data = getSheetData(SHEET_NAMES.ATTENDANCE);
  const monthData = data.filter(r =>
    String(r['従業員ID']) === String(employeeId) && String(r['日付']).startsWith(month)
  );

  if (monthData.length === 0) {
    return `📊 ${employeeName}さんの${month}の出勤履歴\n\nデータがありません。`;
  }

  let totalWork = 0, totalOvertime = 0, totalNight = 0;
  monthData.forEach(r => {
    totalWork += Number(r['実働時間(分)']) || 0;
    totalOvertime += Number(r['残業時間(分)']) || 0;
    totalNight += Number(r['深夜時間(分)']) || 0;
  });

  let msg = `📊 ${employeeName}さんの${month}の出勤履歴\n\n`;
  msg += `出勤日数: ${monthData.length}日\n`;
  msg += `総実働: ${minutesToHM(totalWork)}\n`;
  msg += `総残業: ${minutesToHM(totalOvertime)}\n`;
  if (totalNight > 0) msg += `深夜: ${minutesToHM(totalNight)}\n`;

  // 直近5日分の詳細
  const recent = monthData.slice(-5);
  msg += '\n--- 直近の勤務 ---';
  recent.forEach(r => {
    const d = String(r['日付']).slice(5); // MM/DD
    const ci = r['出勤時刻'] || '--:--';
    const co = r['退勤時刻'] || '--:--';
    const work = r['実働時間(分)'] ? minutesToHM(r['実働時間(分)']) : '-';
    msg += `\n${d} ${ci}〜${co} (${work})`;
  });

  return msg;
}

// ============================================================
// LINE Messaging API
// ============================================================

function handleLineWebhook(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const events = body.events || [];

    events.forEach(event => {
      if (event.type !== 'message' || event.message.type !== 'text') return;

      const userId = event.source.userId;
      const text = event.message.text;
      const replyToken = event.replyToken;

      // メッセージ処理
      const reply = processMessage('line', userId, text);

      // LINE Reply API で応答
      replyLine(replyToken, reply);
    });
  } catch (err) {
    console.error('LINE Webhook Error:', err);
  }
  return ContentService.createTextOutput('OK');
}

function replyLine(replyToken, text) {
  const LINE_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!LINE_TOKEN) {
    console.error('LINE_CHANNEL_ACCESS_TOKEN が未設定です');
    return;
  }

  const payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: String(text) }]
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// === LINE署名検証 ===
function verifyLineSignature(e) {
  const SECRET = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_SECRET');
  if (!SECRET) return true; // 未設定時はスキップ（開発用）

  const signature = e.parameter['x-line-signature'] ||
    (e.headers && e.headers['X-Line-Signature']) || '';
  const body = e.postData.contents;

  const hmac = Utilities.computeHmacSha256Signature(body, SECRET);
  const expected = Utilities.base64Encode(hmac);

  return signature === expected;
}

// ============================================================
// WhatsApp Business API (Meta Cloud API)
// ============================================================

function handleWhatsAppWebhook(e) {
  try {
    // Webhook検証（GET相当のverifyをdoPostで受ける場合）
    if (e.parameter['hub.mode'] === 'subscribe') {
      const verifyToken = PropertiesService.getScriptProperties().getProperty('WHATSAPP_VERIFY_TOKEN');
      if (e.parameter['hub.verify_token'] === verifyToken) {
        return ContentService.createTextOutput(e.parameter['hub.challenge']);
      }
      return ContentService.createTextOutput('Forbidden').setResponseCode(403);
    }

    const body = JSON.parse(e.postData.contents);
    const entries = body.entry || [];

    entries.forEach(entry => {
      const changes = entry.changes || [];
      changes.forEach(change => {
        if (change.field !== 'messages') return;
        const messages = change.value.messages || [];
        messages.forEach(msg => {
          if (msg.type !== 'text') return;

          const phone = msg.from; // 電話番号
          const text = msg.text.body;

          const reply = processMessage('whatsapp', phone, text);
          sendWhatsApp(phone, reply);
        });
      });
    });
  } catch (err) {
    console.error('WhatsApp Webhook Error:', err);
  }
  return ContentService.createTextOutput('OK');
}

function sendWhatsApp(to, text) {
  const TOKEN = PropertiesService.getScriptProperties().getProperty('WHATSAPP_ACCESS_TOKEN');
  const PHONE_ID = PropertiesService.getScriptProperties().getProperty('WHATSAPP_PHONE_NUMBER_ID');
  if (!TOKEN || !PHONE_ID) {
    console.error('WhatsApp API設定が未完了です');
    return;
  }

  const payload = {
    messaging_product: 'whatsapp',
    to: to,
    type: 'text',
    text: { body: String(text) }
  };

  UrlFetchApp.fetch(`https://graph.facebook.com/v18.0/${PHONE_ID}/messages`, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

// ============================================================
// WeChat Official Account API
// ============================================================

function handleWeChatWebhook(e) {
  try {
    // サーバー検証（初回設定時）
    if (e.parameter.echostr) {
      if (verifyWeChatSignature(e)) {
        return ContentService.createTextOutput(e.parameter.echostr);
      }
      return ContentService.createTextOutput('');
    }

    // メッセージ処理
    const xml = e.postData.contents;
    const parsed = parseWeChatXml(xml);
    if (!parsed || parsed.MsgType !== 'text') {
      return ContentService.createTextOutput('');
    }

    const openId = parsed.FromUserName;
    const text = parsed.Content;
    const toUser = parsed.ToUserName;

    const reply = processMessage('wechat', openId, text);

    // XMLパッシブリプライ
    const replyXml = buildWeChatReply(openId, toUser, reply);
    return ContentService.createTextOutput(replyXml)
      .setMimeType(ContentService.MimeType.XML);
  } catch (err) {
    console.error('WeChat Webhook Error:', err);
    return ContentService.createTextOutput('');
  }
}

// === WeChat署名検証 ===
function verifyWeChatSignature(e) {
  const TOKEN = PropertiesService.getScriptProperties().getProperty('WECHAT_TOKEN');
  if (!TOKEN) return true; // 未設定時はスキップ

  const signature = e.parameter.signature;
  const timestamp = e.parameter.timestamp;
  const nonce = e.parameter.nonce;

  const arr = [TOKEN, timestamp, nonce].sort();
  const str = arr.join('');
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, str)
    .map(b => ('0' + ((b < 0 ? b + 256 : b).toString(16))).slice(-2)).join('');

  return hash === signature;
}

// === WeChat XMLパース（簡易） ===
function parseWeChatXml(xml) {
  if (!xml) return null;
  const result = {};
  const tags = ['ToUserName', 'FromUserName', 'MsgType', 'Content', 'CreateTime'];
  tags.forEach(tag => {
    const match = xml.match(new RegExp(`<${tag}><\\!\\[CDATA\\[(.+?)\\]\\]></${tag}>`)) ||
                  xml.match(new RegExp(`<${tag}>(\\d+)</${tag}>`));
    if (match) result[tag] = match[1];
  });
  return result;
}

// === WeChat XMLリプライ組み立て ===
function buildWeChatReply(toUser, fromUser, content) {
  const timestamp = Math.floor(new Date().getTime() / 1000);
  return `<xml>
<ToUserName><![CDATA[${toUser}]]></ToUserName>
<FromUserName><![CDATA[${fromUser}]]></FromUserName>
<CreateTime>${timestamp}</CreateTime>
<MsgType><![CDATA[text]]></MsgType>
<Content><![CDATA[${content}]]></Content>
</xml>`;
}

// ============================================================
// doGet Webhook検証（WhatsApp用）
// ============================================================

// WhatsAppのWebhook検証はGETで来るので、Code.gsのdoGetに追加するか
// 別途URLパラメータで処理する。
// GASのWebアプリではdoGetが1つしか定義できないため、
// URLパラメータで分岐する。
// 例: ?mode=webhook&source=whatsapp&hub.mode=subscribe&...
