/**
 * 経理入力システム v2.0 - メインエントリーポイント
 *
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. 全 .gs ファイルと .html ファイルをコピー
 * 4. GASエディタで initializeSystem() を実行
 * 5. デプロイ → ウェブアプリ → デプロイ
 * 6. デプロイURLにアクセスしてログイン
 *    初期ID: admin / 初期PW: admin123
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('経理入力システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  // Phase 5: メッセージアプリWebhook受信
  try {
    const source = e.parameter.source;
    switch (source) {
      case 'line':
        return handleLineWebhook(e);
      case 'wechat':
        return handleWeChatWebhook(e);
      case 'whatsapp':
        return handleWhatsAppWebhook(e);
      default:
        return ContentService.createTextOutput('OK');
    }
  } catch (err) {
    console.error('doPost error:', err);
    return ContentService.createTextOutput('OK');
  }
}

// Phase 5 スタブ
function handleLineWebhook(e) { return ContentService.createTextOutput('OK'); }
function handleWeChatWebhook(e) { return ContentService.createTextOutput('OK'); }
function handleWhatsAppWebhook(e) { return ContentService.createTextOutput('OK'); }

// HTML include ヘルパー
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}