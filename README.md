# Sales Knowledge Hub v3 — セットアップ・利用ガイド

製造業向け **OCR・BOM・CRM統合型 AI営業ナレッジベース**  
Google Apps Script (GAS) + Google Spreadsheet だけで完結します。

---

## 新機能 (v3 追加点)

| 機能 | 説明 |
|------|------|
| **① OCR・ナレッジ化** | PDF・写真・画像をGeminiがOCRして自動要約・カテゴリ分類・タグ付け。BOM/見積書はデータとして自動抽出 |
| **② BOM管理** | 部品マスタ→基板マスタ→機種マスタの3階層BOM構造。見積書マスタと紐付けて自動展開 |
| **③ CRM** | 会社・連絡先・案件・活動履歴の管理。フェーズ別パイプライン表示。商談議事録と自動連携 |

---

## ファイル構成

```
Code.gs          - バックエンドすべて（OCR / BOM / CRM / AI / DB）
Index.html       - メインUI（GAS Webアプリ）
CSS.html         - スタイルシート（GAS include方式）
JavaScript.html  - フロントエンドJS（GAS include方式）
```

---

## セットアップ手順

### 1. スプレッドシートとApps Scriptの準備

1. Google スプレッドシートを新規作成
2. `拡張機能 > Apps Script` を開く
3. 以下4ファイルを作成してコードを貼り付け：
   - `Code.gs`（既存ファイルに貼り付け）
   - `Index.html`（新規HTMLファイル）
   - `CSS.html`（新規HTMLファイル）
   - `JavaScript.html`（新規HTMLファイル）

### 2. 拡張サービスの有効化

GASエディタ左の **「サービス (＋)」** から以下を追加：
- **Drive API** → OCR処理に必須

### 3. スクリプトプロパティの設定

GASエディタ `⚙ プロジェクトの設定 > スクリプトプロパティ` に以下を追加：

| キー | 値 | 必須 |
|------|----|------|
| `GEMINI_API_KEY` | Gemini API キー | ✅ |
| `KNOWLEDGE_FOLDER_ID` | OCR同期対象フォルダのDrive ID | ✅ |
| `CHAT_WEBHOOK_URL` | Google Chat Webhook URL | 任意 |
| `AUDIO_GEN_WEBHOOK_URL` | n8n 音声生成Webhook | 任意 |

### 4. データベース初期化

1. GASエディタで `setupDatabase` 関数を手動実行
   - またはWebアプリ内の `⚙ DB初期化` ボタンをクリック
2. 以下のシートが自動作成されます：

**ナレッジ系**
- `sources`（OCRナレッジ）, `meetings`（商談議事録）, `notes`（メモ）, `qalogs`（AIログ）, `edges`（グラフエッジ）

**BOM系**
- `parts`（部品マスタ）, `boards`（基板マスタ）, `board_parts`（基板-部品）
- `models`（機種マスタ）, `model_boards`（機種-基板）, `board_files`（ファイルリンク）
- `quotes`（見積書）, `quote_items`（見積明細）

**CRM系**
- `crm_companies`（会社）, `crm_contacts`（連絡先）
- `crm_deals`（案件）, `crm_activities`（活動履歴）

### 5. デプロイ

1. `デプロイ > 新しいデプロイ` をクリック
2. 種類: **ウェブアプリ**
3. アクセス権限: 組織内ユーザー または 全員
4. 発行されたURLにアクセス

---

## 利用ガイド

### ① OCR・ナレッジ化タブ

**Driveフォルダ一括同期:**
- `KNOWLEDGE_FOLDER_ID` フォルダ内のPDF・画像・Google Docsを自動スキャン
- GeminiがOCRテキストを要約・カテゴリ分類・タグ付け
- カテゴリ「BOM」「見積書」のファイルは部品データを自動抽出

**単一ファイル指定:**
- Drive ファイルIDを入力して単体でOCR・ナレッジ化

### ② BOM管理タブ

**登録の流れ:**
```
部品マスタ登録 → 基板マスタ登録 → 機種マスタ登録
     ↓                ↓
  (GASで手動紐付け)  (GASで手動紐付け)
```

> 現在のUIでは各マスタの登録のみ対応。基板-部品の紐付けはGASコード `saveBoardPart()` をAPIで呼び出すか、シートに直接入力してください。

**BOMツリー表示:**
- 機種IDを入力すると、機種→基板→部品の階層をツリー表示
- 各部品の単価・数量から原価積算が可能

**見積書作成:**
- 顧客情報・件名を入力して見積書を作成
- 見積書に機種IDを明細として追加後、BOM自動展開（`buildQuoteBOM`）で部品まで展開

**ファイルスキャン:**
- フォルダID + 命名規則で基板ドキュメントを自動整理
- 例: `KBC-001_部品表.pdf`, `KBC-001_BOM.xlsx`

### ③ CRMタブ

**フェーズ管理:**
`アプローチ → 提案 → 見積提示 → 交渉 → 受注 / 失注`

**活動記録:**
- 訪問・電話・メール・オンライン商談などを記録
- 活動内容は自動的に「商談議事録」にも連携されてAIナレッジに追加

**ダッシュボード:**
- フェーズ別パイプライン棒グラフ
- 担当別案件数・金額
- 最近の活動タイムライン

### AI検索・分析タブ

| モード | 用途 |
|--------|------|
| ナレッジ検索 | 社内データに基づいて質問に回答 |
| デジタルクローン | 過去の商談・ナレッジを自分の記憶として代筆・判断 |
| ロープレ（厳格顧客） | 購買部長として鋭い突っ込みを入れてくれる練習相手 |
| BOM分析 | 部品・基板・機種情報を踏まえた原価・調達分析 |
| CRM分析 | 案件・活動から次のアクションをアドバイス |

---

## 外部API連携 (n8n / スマホ)

デプロイしたURLに対してPOSTリクエストを送ることで外部から操作できます。

```bash
curl -X POST "https://script.google.com/macros/s/YOUR_DEPLOY_ID/exec" \
  -H "Content-Type: application/json" \
  -d '{"action":"saveMeeting","payload":{"client":"A社","summary":"価格交渉あり","fulltext":"...","nextAction":"見積再提出","salesRep":"田中"}}'
```

利用可能なアクション:
- `saveMeeting` / `saveModel` / `saveBoard` / `savePart` / `saveQuote`
- `saveCRMContact` / `saveCRMDeal` / `saveCRMActivity`
- `syncDrive` / `importPDF`
- `processQuery` (AI問い合わせ)
- `getBOMTree` (BOMツリー取得)

---

## 注意事項

- OCR処理はGASの実行時間制限（6分）があります。大量ファイルは分割同期してください
- Gemini API の呼び出し回数制限にご注意ください（無料枠: 15 RPM）
- `Drive API` (拡張サービス) が有効でないとOCRが失敗します
