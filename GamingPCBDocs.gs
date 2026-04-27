/**
 * GamingPCBDocs.gs
 * 遊技機基板 ドキュメント管理モジュール
 *
 * 要件:
 *  1. documents シート初期化（setupGamingDocDatabase）
 *  2. Drive フォルダ内ファイルを機種/基板種類/ドキュメント種別へ自動分類（classifyDriveFiles）
 *  3. Web サイト URL を入力 → タイトル取得 → Gemini 分類 → 保存（saveWebURLAsDocument）
 *  4. ドキュメント一覧取得（getDocuments）/ 手動分類更新（updateDocumentCategory）
 */

// =====================================================================
// 定数定義
// =====================================================================

const GAMING_PCB = {
  SHEET_NAME: 'documents',
  HEADERS: ['ID', 'ファイル名/タイトル', 'URL', '機種名', '基板種類', 'ドキュメント種別', '登録日時'],

  /** 対象とする基板種類 */
  BOARD_TYPES: [
    '枠制御基板',
    '主基板',
    '液晶制御基板',
    '液晶基板',
    'サブ基板',
  ],

  /** ドキュメント種別（全基板共通） */
  DOC_TYPES: [
    '基板設計仕様書',
    '製品仕様書',
    '構成表',
    '部品表',
    '見積書',
    '注文書',
    '納入仕様書',
    'その他書類',
    '売上数',
  ],

  /**
   * キーワードベースの分類ルール
   */
  CLASSIFICATION_RULES: [
    // ── 基板種類ルール ──────────────────────────────
    { field: 'boardType', label: '枠制御基板',   keywords: ['枠制御', '枠基板', 'frame_ctrl', 'frame_board'] },
    { field: 'boardType', label: '主基板',        keywords: ['主基板', 'main_board', 'mainboard', 'main基板'] },
    { field: 'boardType', label: '液晶制御基板',  keywords: ['液晶制御', 'lcd_ctrl', 'lcd_control'] },
    { field: 'boardType', label: '液晶基板',      keywords: ['液晶基板', 'lcd_board', 'lcdboard'] },
    { field: 'boardType', label: 'サブ基板',      keywords: ['サブ基板', 'sub_board', 'subboard', 'sub基板'] },

    // ── ドキュメント種別ルール ──────────────────────
    { field: 'docType', label: '基板設計仕様書', keywords: ['基板設計仕様', '基板仕様書', 'board_spec', 'pcb_spec', '回路設計仕様'] },
    { field: 'docType', label: '製品仕様書',     keywords: ['製品仕様', 'product_spec', '機種仕様', 'お客様支給', 'model_spec'] },
    { field: 'docType', label: '構成表',         keywords: ['構成表', '構成リスト', 'bom_list', '構成図', 'configuration'] },
    { field: 'docType', label: '部品表',         keywords: ['部品表', 'BOM', 'parts_list', 'bill_of_material', '部品リスト'] },
    { field: 'docType', label: '見積書',         keywords: ['見積', '見積書', 'quote', 'quotation', 'estimate'] },
    { field: 'docType', label: '注文書',         keywords: ['注文書', '発注書', 'purchase_order', '注文'] },
    { field: 'docType', label: '納入仕様書',     keywords: ['納入仕様', '納品仕様', 'delivery_spec', '納入条件'] },
    { field: 'docType', label: '売上数',         keywords: ['売上', '販売数', '出荷数', 'sales', 'shipment'] },
  ],
};

// =====================================================================
// 1. データベース初期化
// =====================================================================

/**
 * documents シートを作成または検証し、ヘッダーを設定する。
 * 既存の setupDatabase() の末尾に setupGamingDocDatabase(); を追記してください。
 */
function setupGamingDocDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GAMING_PCB.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(GAMING_PCB.SHEET_NAME);
  }

  const existing = sheet.getRange(1, 1, 1, GAMING_PCB.HEADERS.length).getValues()[0];
  const needsHeader = GAMING_PCB.HEADERS.some((h, i) => existing[i] !== h);
  if (needsHeader) {
    sheet.getRange(1, 1, 1, GAMING_PCB.HEADERS.length).setValues([GAMING_PCB.HEADERS]);
    sheet.getRange(1, 1, 1, GAMING_PCB.HEADERS.length)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  Logger.log('documents シート初期化完了');
  return { success: true, message: 'documents シートを初期化しました' };
}

// =====================================================================
// 内部ユーティリティ
// =====================================================================

/** documents シートに1行追記してIDを返す */
function _appendDocument(fileName, url, modelName, boardType, docType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GAMING_PCB.SHEET_NAME);
  if (!sheet) {
    setupGamingDocDatabase();
    sheet = ss.getSheetByName(GAMING_PCB.SHEET_NAME);
  }

  // Code.gs の generateId() と同様に UUID を使用（タイムスタンプ方式はミリ秒単位で重複リスクあり）
  const id  = 'DOC-' + Utilities.getUuid();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

  sheet.appendRow([id, fileName || '', url || '', modelName || '', boardType || '', docType || '', now]);
  return id;
}

/**
 * テキスト（ファイル名や抽出本文）からキーワードルールで分類する。
 * @returns {{ boardType: string, docType: string }}
 */
function _classifyByKeywords(text) {
  const lower = String(text || '').toLowerCase();
  let boardType = '';
  let docType   = '';

  for (const rule of GAMING_PCB.CLASSIFICATION_RULES) {
    for (const kw of rule.keywords) {
      if (lower.includes(kw.toLowerCase())) {
        if (rule.field === 'boardType' && !boardType) boardType = rule.label;
        if (rule.field === 'docType'   && !docType)   docType   = rule.label;
      }
    }
    if (boardType && docType) break; // 両方確定したら早期終了
  }

  return { boardType, docType };
}

/**
 * Gemini API を使って分類を推論する（キーワードで未確定の場合に利用）。
 * @param {string} title  ファイル名やページタイトル
 * @param {string} body   抽出テキスト（省略可）
 * @returns {{ modelName: string, boardType: string, docType: string }}
 */
function _classifyByGemini(title, body) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { modelName: '', boardType: '', docType: '' };

  const boardList = GAMING_PCB.BOARD_TYPES.join('、');
  const docList   = GAMING_PCB.DOC_TYPES.join('、');

  const prompt = `あなたは遊技機基板メーカーの書類分類AIです。
以下の情報をもとに、JSON形式のみで分類してください。

タイトル: ${title}
本文（先頭500文字）: ${String(body || '').slice(0, 500)}

分類ルール:
- modelName: 機種名（例: XX-1000、〇〇遊技機など。不明なら空文字）
- boardType: 基板種類（${boardList}のいずれか。不明なら空文字）
- docType: ドキュメント種別（${docList}のいずれか。不明なら空文字）

必ずJSON形式のみで返してください。例:
{"modelName":"XX-1000","boardType":"枠制御基板","docType":"基板設計仕様書"}`;

  try {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    const res = UrlFetchApp.fetch(endpoint, {
      method:      'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents:         [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 200 },
      }),
      muteHttpExceptions: true,
    });

    const json = JSON.parse(res.getContentText());
    const raw  = json?.candidates?.[0]?.content?.parts?.[0]?.text || '';

    // レスポンスから最初の JSON オブジェクトを抽出（```json ... ``` ラッパーや前後テキストに対応）
    const jsonMatch = raw.match(/\{[\s\S]*?\}/);
    if (!jsonMatch) return { modelName: '', boardType: '', docType: '' };

    return JSON.parse(jsonMatch[0]);
  } catch (e) {
    Logger.log('Gemini分類エラー: ' + e.message);
    return { modelName: '', boardType: '', docType: '' };
  }
}

// =====================================================================
// 2. Drive ファイルの自動分類と転記
// =====================================================================

/**
 * 指定した Drive フォルダ内のファイルを再帰巡回し、
 * ファイル名 → キーワードルール → Gemini の順で分類して documents シートに転記する。
 *
 * @param {string} folderId  対象 Google Drive フォルダ ID（省略時はマイドライブ直下）
 * @returns {{ success: boolean, added: number, skipped: number }}
 */
function classifyDriveFiles(folderId) {
  setupGamingDocDatabase(); // シート未作成の場合に備えて初期化

  const folder = folderId
    ? DriveApp.getFolderById(folderId)
    : DriveApp.getRootFolder();

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GAMING_PCB.SHEET_NAME);

  // 既存URLをSetに格納（O(1)検索）
  const existingUrls = new Set(
    sheet && sheet.getLastRow() > 1
      ? sheet.getDataRange().getValues().slice(1).map(r => String(r[2])).filter(Boolean)
      : []
  );

  let added   = 0;
  let skipped = 0;

  // 再帰巡回（クロージャで added/skipped/existingUrls を共有）
  function processFolder(f) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const url  = file.getUrl();

      if (existingUrls.has(url)) {
        skipped++;
        continue;
      }

      const name = file.getName();
      let { boardType, docType } = _classifyByKeywords(name);

      // キーワードで未確定の項目を Gemini で補完
      if (!boardType || !docType) {
        const g = _classifyByGemini(name, '');
        if (!boardType) boardType = g.boardType || 'その他書類';
        if (!docType)   docType   = g.docType   || 'その他書類';
        _appendDocument(name, url, g.modelName || '', boardType, docType);
      } else {
        _appendDocument(name, url, '', boardType, docType);
      }

      existingUrls.add(url); // 同セッション内での重複防止
      added++;
    }

    const subs = f.getFolders();
    while (subs.hasNext()) processFolder(subs.next());
  }

  processFolder(folder);
  return { success: true, added, skipped };
}

// =====================================================================
// 3. Web サイト URL からの自動保存
// =====================================================================

/**
 * URL を受け取り、ページ情報を取得 → Gemini で分類 → documents シートに保存する。
 * フロントエンドの google.script.run から直接呼び出す。
 *
 * @param {string} url       保存したい Web ページの URL
 * @param {string} modelName 機種名（ユーザー入力。省略可）
 * @returns {{ success: boolean, id?: string, title?: string, modelName?: string, boardType?: string, docType?: string, message?: string }}
 */
function saveWebURLAsDocument(url, modelName) {
  if (!url) return { success: false, message: 'URLを入力してください' };

  let title    = url; // フェッチ失敗時のフォールバック
  let bodyText = '';

  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const html = res.getContentText('UTF-8');

    // <title> 抽出（改行・空白を含むタイトルに対応するため [\s\S]+? を使用）
    const titleMatch = html.match(/<title[^>]*>([\s\S]+?)<\/title>/i);
    if (titleMatch) title = titleMatch[1].replace(/\s+/g, ' ').trim();

    // description メタタグ抽出（name/property と content の順序不問に対応）
    const metaDesc = html.match(
      /<meta[^>]+(?:name|property)=["'](?:description|og:description)["'][^>]+content=["']([\s\S]+?)["']/i
    ) || html.match(
      /<meta[^>]+content=["']([\s\S]+?)["'][^>]+(?:name|property)=["'](?:description|og:description)["']/i
    );
    if (metaDesc) bodyText += metaDesc[1].trim() + ' ';

    // スクリプト・スタイルを除去してプレーンテキスト化（先頭1000文字）
    const plainText = html
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    bodyText += plainText.slice(0, 1000);

  } catch (e) {
    Logger.log('URLフェッチエラー: ' + e.message);
    // フェッチ失敗時はタイトル=URL・空ボディのまま Gemini/キーワード分類を続行
  }

  // Gemini で一次分類
  const gemini         = _classifyByGemini(title, bodyText);
  const finalModel     = modelName || gemini.modelName || '';
  let   finalBoardType = gemini.boardType || '';
  let   finalDocType   = gemini.docType   || '';

  // Gemini で未確定の項目をキーワードで補完
  if (!finalBoardType || !finalDocType) {
    const kw = _classifyByKeywords(title + ' ' + bodyText);
    if (!finalBoardType) finalBoardType = kw.boardType || 'その他書類';
    if (!finalDocType)   finalDocType   = kw.docType   || 'その他書類';
  }

  const id = _appendDocument(title, url, finalModel, finalBoardType, finalDocType);

  return {
    success:   true,
    id,
    title,
    modelName: finalModel,
    boardType: finalBoardType,
    docType:   finalDocType,
  };
}

// =====================================================================
// 4. ドキュメント一覧取得 / 更新
// =====================================================================

/**
 * documents シートの全データを返す。
 * フロントエンドの google.script.run から直接呼び出す。
 *
 * @param {Object} filter  { modelName?, boardType?, docType? } — 全て省略で全件
 * @returns {Object[]}  ドキュメントオブジェクトの配列
 */
function getDocuments(filter) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GAMING_PCB.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  let docs = sheet.getDataRange().getValues().slice(1)
    .map(r => ({
      id:        String(r[0] || ''),
      fileName:  String(r[1] || ''),
      url:       String(r[2] || ''),
      modelName: String(r[3] || ''),
      boardType: String(r[4] || ''),
      docType:   String(r[5] || ''),
      createdAt: String(r[6] || ''),
    }))
    .filter(d => d.id); // 空行除外

  if (filter) {
    if (filter.modelName) docs = docs.filter(d => d.modelName === filter.modelName);
    if (filter.boardType) docs = docs.filter(d => d.boardType === filter.boardType);
    if (filter.docType)   docs = docs.filter(d => d.docType   === filter.docType);
  }

  return docs;
}

/**
 * 指定 ID のドキュメントの分類情報を手動で更新する。
 *
 * @param {string} id        DOC-XXX 形式の ID
 * @param {string} boardType 新しい基板種類
 * @param {string} docType   新しいドキュメント種別
 * @param {string} modelName 新しい機種名（省略可）
 * @returns {{ success: boolean, message?: string }}
 */
function updateDocumentCategory(id, boardType, docType, modelName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GAMING_PCB.SHEET_NAME);
  if (!sheet) return { success: false, message: 'シートが存在しません' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      if (modelName !== undefined) sheet.getRange(i + 1, 4).setValue(modelName);
      if (boardType !== undefined) sheet.getRange(i + 1, 5).setValue(boardType);
      if (docType   !== undefined) sheet.getRange(i + 1, 6).setValue(docType);
      return { success: true };
    }
  }
  return { success: false, message: 'IDが見つかりません: ' + id };
}

/**
 * フロントエンドで使うマスターデータ（基板種類・ドキュメント種別リスト）を返す。
 */
function getGamingPCBMasters() {
  return {
    boardTypes: GAMING_PCB.BOARD_TYPES,
    docTypes:   GAMING_PCB.DOC_TYPES,
  };
}
