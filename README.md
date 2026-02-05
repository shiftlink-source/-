# 📊 経理入力システム v2.0

Google Apps Script + Google Sheets で構築する中小企業向け経理・勤怠管理システム

## 🎯 概要

スプレッドシートをバックエンドDBとして活用し、GAS Webアプリで以下の機能を提供します：

- **経理入力**: 小口現金管理、銀行/クレカCSV取り込み、AI仕訳推定
- **マスタ管理**: 勘定科目・取引先・口座の一元管理
- **出退勤管理**: LINE/WhatsApp/WeChat連携の打刻システム
- **請求書・支払明細**: インボイス制度対応のPDF自動生成
- **入金消込**: 売掛金・買掛金の突合・消込処理

## 📁 プロジェクト構成

```
src/
├── Code.gs              # メインエントリーポイント（doGet/doPost）
├── Config.gs            # 定数・シート名・デフォルトデータ定義
├── Utils.gs             # ユーティリティ関数
├── Auth.gs              # 認証・セッション管理
├── SheetSetup.gs        # シート初期化・デフォルトデータ投入
├── MasterService.gs     # マスタCRUD（勘定科目・取引先・口座・設定）
├── PettyCash.gs         # 小口現金管理
├── OperationLogger.gs   # 操作ログ
├── index.html           # メインUI（SPAシェル）
├── Stylesheet.html      # CSS
└── JavaScript.html      # クライアントサイドJS
```

## 🚀 セットアップ

### 前提条件
- Google アカウント
- Google スプレッドシート

### 手順

#### 1. スプレッドシート作成
[Google Sheets](https://sheets.google.com) で新規スプレッドシートを作成

#### 2. Apps Script プロジェクト作成
メニュー → **拡張機能** → **Apps Script**

#### 3. ファイルをコピー

`src/` 内のファイルをGASエディタに追加します。

**スクリプトファイル（`.gs`）:**
GASエディタで **＋ → スクリプト** から以下を追加：
- `Code` / `Config` / `Utils` / `Auth` / `SheetSetup` / `MasterService` / `PettyCash` / `OperationLogger`

**HTMLファイル（`.html`）:**
GASエディタで **＋ → HTML** から以下を追加：
- `index` / `Stylesheet` / `JavaScript`

> 💡 **Tips:** [clasp](https://github.com/google/clasp) を使えばローカルから直接push可能です（後述）

#### 4. システム初期化
1. GASエディタで `SheetSetup.gs` を開く
2. 関数セレクタで `initializeSystem` を選択
3. **▶ 実行** → 権限承認 → 完了

以下が自動生成されます：
- 全24シート
- デフォルト勘定科目（約50件）
- デフォルト管理者アカウント

#### 5. Webアプリとしてデプロイ
**デプロイ** → **新しいデプロイ** → **ウェブアプリ**

| 設定 | 値 |
|------|-----|
| 次のユーザーとして実行 | 自分 |
| アクセスできるユーザー | 全員（※必要に応じて制限） |

#### 6. 初回ログイン

| 項目 | 値 |
|------|-----|
| ユーザーID | `admin` |
| パスワード | `admin123` |

> ⚠️ **初回ログイン後、必ずパスワードを変更してください**

---

## 🔧 clasp によるローカル開発（推奨）

[clasp](https://github.com/google/clasp) を使えばGitHub管理とGASへのデプロイを連携できます。

```bash
# clasp インストール
npm install -g @google/clasp

# ログイン
clasp login

# プロジェクトとの紐付け（既存のGASプロジェクトID指定）
clasp clone <script-id> --rootDir src

# ローカル編集後にpush
clasp push

# デプロイ
clasp deploy --description "v2.0 Phase 1"
```

### `.clasp.json` の例
```json
{
  "scriptId": "<YOUR_SCRIPT_ID>",
  "rootDir": "src"
}
```

---

## 📅 開発ロードマップ

| Phase | 内容 | 期間目安 | 状態 |
|-------|------|---------|------|
| 1 | 基盤構築（マスタ・認証・ログイン・小口現金） | 1-2週間 | ✅ 完了 |
| 2 | CSVインポート（銀行・クレカ） | 1-2週間 | ⬜ |
| 3 | AI仕訳推定（Gemini API連携） | 1週間 | ⬜ |
| 4 | 仕訳帳＋エクスポート（MF/弥生/freee） | 1-2週間 | ⬜ |
| 5 | 出退勤管理（LINE/WhatsApp/WeChat） | 2-3週間 | ⬜ |
| 6 | 請求書・支払明細（インボイス対応） | 2週間 | ⬜ |
| 7 | 入金消込 | 2週間 | ⬜ |
| 8 | ダッシュボード拡張・仕上げ | 1-2週間 | ⬜ |

---

## 🏗️ アーキテクチャ

```
[ブラウザ]
  │  SPA（HTML/CSS/JS）
  │  google.script.run で API 呼び出し
  │
[GAS Webアプリ]
  ├── doGet()  → HTML配信
  ├── doPost() → Webhook受信（LINE/WhatsApp/WeChat）
  ├── Auth     → セッション管理（CacheService）
  ├── Services → ビジネスロジック
  └── Utils    → 共通関数
  │
[Google スプレッドシート]
  └── 24シート（マスタ・取引・勤怠・帳票・債権債務・ログ）
```

---

## 📋 スプレッドシート構成（全24シート）

**マスタ:** 設定 / 勘定科目マスタ / 取引先マスタ / 口座マスタ / 従業員マスタ / 勤怠設定 / 認証マスタ

**取引データ:** 小口現金 / GMOあおぞら / 住信SBI / みずほ / クレカ_JCB / クレカ_VISA

**勤怠データ:** 出退勤データ / 打刻ログ

**帳票データ:** 請求書マスタ / 請求書明細 / 支払明細マスタ / 支払明細_明細

**債権債務:** 売掛金管理 / 買掛金管理 / 消込明細

**統合・ログ:** 仕訳帳 / AI学習データ / 操作ログ

---

## ⚠️ 制約事項

| 項目 | 制限 | 対策 |
|------|------|------|
| GAS実行時間 | 最大6分/回 | 処理の分割 |
| セッション | 最大6時間（CacheService） | 自動延長あり |
| スプレッドシート行数 | 最大1,000万セル | 年度ごとにファイル分割 |
| 同時書き込み | ロック機構なし | 競合時はリトライ |
| Webhookレスポンス | ContentServiceのみ | 非同期応答 |

---

## 📄 ライセンス

MIT License

---

## 🤝 コントリビューション

Issue・Pull Request 歓迎です。

---

*モデルケース: 株式会社レイズ*
