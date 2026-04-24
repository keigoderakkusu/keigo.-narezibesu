# Drive自動分類システム 組み込み手順

## 追加されるファイル

| ファイル名 | 追記先 | 役割 |
|---|---|---|
| `DriveClassifier.gs` | GASに新規ファイルとして追加 | バックエンド：Drive自動スキャン・分類ロジック |
| `INDEX_ADDITION.html` | `Index.html` に組み込む | UI：機種ツリータブ＆スキャンモーダル |
| `JS_ADDITION.html` | `JavaScript.html` に組み込む | フロント：ツリー描画・モーダル制御 |

---

## Step 1: GAS に DriveClassifier.gs を追加

1. GASエディタを開く
2. 左側の「+」ボタン → 「スクリプト」を選択
3. ファイル名を `DriveClassifier` にする
4. `DriveClassifier.gs` の内容を全てコピーして貼り付け
5. 保存（Ctrl+S）

---

## Step 2: Index.html にタブとモーダルを追加

### 2-1. ナビゲーションに追加

```html
<!-- navbarNav の <ul> 内、最後の </ul> の直前に追記 -->
<li class="nav-item">
  <a class="nav-link" id="nav-modeltree" href="#" onclick="showTab('modeltree')">
    <i class="bi bi-diagram-3 text-info me-1"></i>機種・ドキュメント管理
  </a>
</li>
```

### 2-2. タブ本体を追加

`INDEX_ADDITION.html` の中の `<div id="tab-modeltree" ...>` から
最初の `</div>` （タブ終わり）までを、
他のタブdiv（例: `tab-notebooklm`）の **直後** にコピー

### 2-3. モーダルを追加

`INDEX_ADDITION.html` の中の `<div class="modal fade" id="modal-scan" ...>` から
末尾の `</div>` までを、
他のモーダル（例: `modal-goal`）の **直後** にコピー

---

## Step 3: JavaScript.html に JS を追加

`JS_ADDITION.html` の内容（`<script>` タグの中身）を、
`JavaScript.html` の `</script>` の **直前** にコピー

---

## Step 4: シートの自動生成を確認

1. GASエディタで `ensureClassifierSheets` を手動実行
   （または初回スキャン時に自動生成されます）
2. スプレッドシートに以下のシートが追加されることを確認：
   - `models`（機種マスタ）
   - `boards`（基板マスタ）
   - `parts`（部品マスタ）
   - `quotes`（見積書台帳）
   - `doc_hierarchy`（ドキュメント階層）

---

## 使い方

### Driveフォルダのスキャン

1. 「機種・ドキュメント管理」タブを開く
2. 「Driveフォルダを取込む」ボタンをクリック
3. **機種名**と **DriveフォルダのID** を入力
   - フォルダIDは Drive URL の末尾部分
   - `https://drive.google.com/drive/folders/` **← ここがID**
4. 「検証」ボタンでフォルダの存在確認
5. 「スキャン開始」→ 自動分類が実行される

### 自動分類ルール

| ファイル名キーワード | 分類先 |
|---|---|
| 製品仕様, spec, 仕様書 | 📋 機種製品仕様書 |
| 基板設計, 回路設計, PCB, schematic | 🔧 基板設計仕様書 |
| BOM, 構成表, 部品表, bill_of_material | 📦 構成表(BOM) |
| 見積, quote, QUOTE, estimate | 💴 見積書 |
| 図面, drawing, DWG, CAD | 📐 図面 |
| マニュアル, manual, 取扱説明 | 📖 マニュアル |
| 試験, テスト, test, 検査 | 🧪 試験・検査書 |

- **サブフォルダ名に「基板」「PCB」「board」が含まれる場合**: 基板ノードとして自動認識
- **Excelファイル (.xlsx)**: 構成表(BOM)として自動分類
- **上記にマッチしないファイル**: 「未分類」として登録 → 後から手動で種別変更可能

### 未分類ファイルの種別変更

1. ツリー上の「⚠️ 未分類」ファイルの「分類する」ボタンをクリック
2. 種別番号を入力して確定
3. ツリーが自動更新される
