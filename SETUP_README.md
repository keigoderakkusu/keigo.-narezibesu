# 4機能 追加手順書

## 追加されるファイル一覧

| ファイル名 | 役割 |
|---|---|
| `Feature_Vector.gs` | ⑪ Gemini Embedding APIによる意味検索エンジン |
| `Feature_Approval.gs` | ⑫ 記事承認ワークフロー（Gmail通知付き） |
| `Feature_Import.gs` | ⑧ スプレッドシート一括インポート（AI自動種別判定） |
| `Feature_KPI.gs` | ⑨ KPIダッシュボード（カレンダー連動・月次レポート） |
| `IndexAddition.html` | 4機能のUIパーツ（Indexに追記） |
| `JS_Features.html` | 4機能のJS（JavaScript.htmlに追記） |

---

## Step 1: GASに4つの.gsファイルを追加

GASエディタで「新しいスクリプトファイル」を4つ作成し、それぞれの内容を貼り付けます。

```
Feature_Vector
Feature_Approval
Feature_Import
Feature_KPI
```

---

## Step 2: DBシートを初期化

GASエディタで以下の関数を順番に手動実行します：

```javascript
setupVectorDB()    // vector_index シートを作成
setupApprovalDB()  // approval_requests, approval_logs シートを作成
setupImportDB()    // import_logs シートを作成
setupKpiDB()       // kpi_targets, kpi_snapshots シートを作成
```

---

## Step 3: Index.html に UIパーツを追記

### 3-1. サイドバーにナビ項目を追加

`IndexAddition.html` 冒頭のコメントアウトされた `sb-item` 4つを、
既存 Index.html の `<div class="sb-section">その他</div>` の **直前** に追記します：

```html
<div class="sb-section" style="margin-top:16px;">高度な機能</div>
<div class="sb-item" id="sb-vector"   onclick="showPage('vector')"><i class="bi bi-search-heart"></i><span>意味検索</span></div>
<div class="sb-item" id="sb-approval" onclick="showPage('approval')"><i class="bi bi-shield-check"></i><span>承認ワークフロー</span><span class="sb-badge" id="sb-approval-cnt">0</span></div>
<div class="sb-item" id="sb-import"   onclick="showPage('import')"><i class="bi bi-cloud-upload"></i><span>一括インポート</span></div>
<div class="sb-item" id="sb-kpi"      onclick="showPage('kpi')"><i class="bi bi-graph-up-arrow"></i><span>KPI ダッシュボード</span></div>
```

### 3-2. ページブロックを追加

`IndexAddition.html` の `[2] ページブロック` セクション（4つの `<div class="page">` ブロック）を、
`</main>` の **直前** にまるごとコピーします。

### 3-3. モーダルを追加

`IndexAddition.html` の `[3] 追加モーダル` セクション（承認モーダル・KPI目標モーダル）を、
既存モーダルの末尾（`</body>` の直前）にコピーします。

---

## Step 4: JavaScript.html に JS を追記

`JS_Features.html` の `<script>` タグの **中身だけ** を、
`JavaScript.html` の既存 `</script>` の **直前** にコピーします。

---

## Step 5: スクリプトプロパティの設定

GASエディタ → 「プロジェクトの設定」→「スクリプトプロパティ」で以下を追加：

| キー | 値 | 用途 |
|---|---|---|
| `GEMINI_API_KEY` | Gemini APIキー | ⑪⑧ 必須 |
| `APPROVER_EMAIL` | 承認者のGmailアドレス | ⑫ 承認通知メール |
| `APP_URL` | WebアプリのURL | ⑫ メール内リンク |

---

## Step 6: 定期トリガーの設定（任意）

GASエディタ → 「トリガー」で以下を設定すると自動化が完了します：

| 関数名 | 頻度 | 用途 |
|---|---|---|
| `buildVectorIndex` | 毎日 午前2時 | ⑪ インデックス自動更新 |
| `sendApprovalReminders` | 毎日 午前9時 | ⑫ 承認忘れリマインダー |
| `sendMonthlyKpiReport` | 毎月1日 午前8時 | ⑨ 月次KPIレポート自動送信 |

---

## 各機能の使い方

### ⑪ 意味検索
1. サイドバー「意味検索」を開く
2. 「インデックスを構築」ボタンを押す（初回のみ、数分かかります）
3. 検索バーに自然な質問文を入力 → 検索
   - 例: 「省エネ基板の設計上の注意点は？」
   - 例: 「A社との商談で出た価格に関する課題」

### ⑫ 承認ワークフロー
- **記事作成者**: KB記事を書いた後、記事詳細パネルの「承認申請」ボタンを押す
- **承認者**: サイドバー「承認ワークフロー」→「承認待ち」タブで確認・承認/差し戻し
- `APPROVER_EMAIL` を設定するとGmailで通知が届きます

### ⑧ 一括インポート
1. サイドバー「一括インポート」を開く
2. 「URLで指定」タブ → スプレッドシートのURLを貼り付け
3. 「プレビュー確認」で内容を確認してから「インポート実行」
   - 種別はAIが自動判定（手動指定も可能）

### ⑨ KPIダッシュボード
1. サイドバー「KPI ダッシュボード」を開く
2. 「目標を設定」ボタンでKPIを登録
   - 例: 商談件数 目標10件/月
3. ダッシュボードがリアルタイムで実績と比較を表示
4. Googleカレンダーの予定（「商談」「訪問」などが含まれるもの）が自動連携
