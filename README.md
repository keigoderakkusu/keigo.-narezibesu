# Sales Knowledge Hub (完全GAS・Google Drive統合版)

製造業の「営業DX」および「セールス・イネーブルメント」を強力に推し進める、完全Google Workspace完結型のAI営業ナレッジベースです。

## 特徴
1. **外部サーバー完全不要**: 全てGASとGoogle Spreadsheetsの上で動作します。
2. **圧倒的なRAG能力 (Gemini 1.5対応)**: 組織内の「商談議事録」「製品マスタ」「Obsidianメモ」だけでなく、**Google DriveにアップロードされたPDFやExcel（ドキュメント）を直接AIコンテキストに読み込ませる**ことができます。

## 初期設定手順

### 1. Google スプレッドシート / GAS の準備
1. 新規Googleスプレッドシートを作成し、`拡張機能 > Apps Script` を開きます。
2. 本リポジトリ内の `Code.gs`, `Index.html`, `JavaScript.html`, `CSS.html` の内容を貼り付けて保存します。

### 2. Google Drive 「ナレッジハブ」フォルダの準備
1. ご自身のGoogle Driveに、AIに読ませたいPDFやドキュメントを入れるフォルダ（例：「AI営業ファイル連携用」）を作成します。
2. フォルダを開いたURL (`https://drive.google.com/drive/folders/〇〇〇〇`) にある、`〇〇〇〇` の部分（フォルダID）をコピーします。

### 3. スクリプトプロパティ（APIキー・フォルダID）の設定
1. GASエディタの歯車マーク（プロジェクトの設定）を開きます。
2. ページ下部の「スクリプトプロパティ」で以下を追加します。
   - `GEMINI_API_KEY`: ご自身のGemini APIキー
   - `KNOWLEDGE_FOLDER_ID`: 先ほどコピーしたGoogle DriveのフォルダID
   - `CHAT_WEBHOOK_URL` (オプション): Google ChatのWebhook URL（業務分析＆チャット通知用）

### 4. 拡張機能「Drive API」の有効化 (OCR用)
1. GASエディタ左側の `サービス (+) ` ボタンをクリックします。
2. APIリストから **Drive API** を探し、「追加」をクリックします。（※これがないとPDFのOCRテキスト抽出が失敗します）

### 5. 初回自動構築とデプロイ
1. エディタ上部の関数実行プルダウンから `setupDatabase` を手動実行し、スプレッドシートにテーブルを自動生成させます。
2. 右上の青いボタン `[デプロイ] > [新しいデプロイ]` をクリックします。
3. アプリの種類に「ウェブアプリ」を選び、対象ユーザーを「全員」、または「組織内ユーザー」としてデプロイします。

## 使い方
1. 発行されたウェブアプリのURLにアクセスします。
2. **「Driveフォルダ同期」タブ**: Google Driveに入れたPDF等をRAG用のナレッジとしてシートに取り込みます。
3. **「AI検索・ロープレ」タブ**: Gemini 1.5 が読み込んだ最新のPDFや製品情報を元に、見積りの相談や営業ロープレを行ってくれます。
