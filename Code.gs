// ==========================================
// Code.gs v3.0 — OCR + BOM + CRM 統合版
// ==========================================

const SCRIPT_PROP_GEMINI_KEY = 'GEMINI_API_KEY';
const SCRIPT_PROP_FOLDER_ID  = 'KNOWLEDGE_FOLDER_ID';
const SCRIPT_PROP_CHAT_WEBHOOK = 'CHAT_WEBHOOK_URL';
const SCRIPT_PROP_AUDIO_WEBHOOK = 'AUDIO_GEN_WEBHOOK_URL';

// ==========================================
// Web App Entry Points
// ==========================================

function doGet(e) {
  const tmpl = HtmlService.createTemplateFromFile('Index');
  return tmpl.evaluate()
    .setTitle('Sales Knowledge Hub v3')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const handlers = {
      'saveMeeting'       : () => saveMeeting(data.payload),
      'processQuery'      : () => processAIQuery(data.payload.query, data.payload.mode || 'search'),
      'saveModel'         : () => saveModel(data.payload),
      'saveBoard'         : () => saveBoard(data.payload),
      'savePart'          : () => savePart(data.payload),
      'saveQuote'         : () => saveQuote(data.payload),
      'saveQuoteItem'     : () => saveQuoteItem(data.payload),
      'saveCRMContact'    : () => saveCRMContact(data.payload),
      'saveCRMActivity'   : () => saveCRMActivity(data.payload),
      'saveCRMDeal'       : () => saveCRMDeal(data.payload),
      'importPDF'         : () => importPDFAndClassify(data.payload.fileId, data.payload.context),
      'importCSV'         : () => importBulkData(data.payload.sheetType, data.payload.headers, data.payload.rows),
      'syncDrive'         : () => syncDriveKnowledge(),
      'saveEdge'          : () => saveEdge(data.payload),
      'autoExtractEdges'  : () => autoExtractEdges(),
      'saveGoal'          : () => saveGoal(data.payload),
      'deleteRow'         : () => deleteRowData(data.payload.sheetName, data.payload.id),
      'updateRow'         : () => updateRowData(data.payload.sheetName, data.payload.id, data.payload.col, data.payload.val),
      'getBOMTree'        : () => getBOMTree(data.payload.modelId),
      'buildQuoteBOM'     : () => buildQuoteBOM(data.payload.quoteId),
      'scanDriveFiles'    : () => scanAndLinkBoardFiles(data.payload.folderId, data.payload.regex),
      'analyzeWorkspace'  : () => analyzeWorkspaceAndPushChat(),
    };
    const fn = handlers[action];
    if (!fn) return _jsonOut({ status: 'error', message: '不明なアクション: ' + action });
    const result = fn();
    return _jsonOut({ status: 'success', result });
  } catch (err) {
    return _jsonOut({ status: 'error', message: err.message });
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function _jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// DB セットアップ (v3 拡張)
// ==========================================

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsConfig = {
    // ----- ナレッジ -----
    'sources'       : ['ID','ファイル名','URL','タイプ','連携日時','OCRテキスト','AI要約','カテゴリ','タグ'],
    'meetings'      : ['ID','登録日時','顧客名','関連基板・機種','内容サマリ','議事録全文','次回アクション','担当営業'],
    'notes'         : ['ID','登録日時','タグ','メモ内容'],
    'qalogs'        : ['日時','ユーザー入力','AI回答'],
    'edges'         : ['ID','Source_ID','Target_ID','Relation_Type','登録日時'],
    'goals'         : ['ID','登録日時','対象期間','目標タイトル','KPI','進捗(%)','評価','関連ID'],

    // ----- BOM マスタ群 -----
    'parts'         : ['部品ID','部品名','部品番号','メーカー','仕様','単価','通貨','リードタイム(日)','在庫数','備考','登録日時'],
    'boards'        : ['基板ID','基板名','リビジョン','説明','親製品ID','製造コスト','標準工数(h)','ステータス','登録日時'],
    'board_parts'   : ['ID','基板ID','部品ID','数量','マウント位置','備考'],
    'models'        : ['機種ID','機種名','型番','カテゴリ','説明','販売価格','原価','マージン率(%)','ステータス','後継機種ID','登録日時'],
    'model_boards'  : ['ID','機種ID','基板ID','数量','役割','備考'],
    'board_files'   : ['基板ID','部品表URL','構成表URL','BOM_URL','仕様書URL','スキャン日時'],

    // ----- 見積書マスタ -----
    'quotes'        : ['見積ID','見積番号','顧客名','顧客ID','件名','作成日','有効期限','合計金額','税率(%)','税込合計','ステータス','担当営業','備考'],
    'quote_items'   : ['ID','見積ID','機種ID','基板ID','部品ID','品名','数量','単価','金額','備考'],

    // ----- CRM -----
    'crm_companies' : ['会社ID','会社名','業種','都道府県','住所','電話','URL','担当営業','ランク','備考','登録日時'],
    'crm_contacts'  : ['連絡先ID','会社ID','氏名','役職','メール','電話','担当営業','登録日時','備考'],
    'crm_deals'     : ['案件ID','会社ID','連絡先ID','件名','フェーズ','金額','確度(%)','受注予定日','担当営業','見積ID','登録日時','備考'],
    'crm_activities': ['活動ID','案件ID','会社ID','連絡先ID','種別','日時','内容','次回アクション','担当営業'],
  };

  const headerColors = {
    'sources'       : '#7c3aed',
    'meetings'      : '#b45309',
    'notes'         : '#15803d',
    'edges'         : '#be123c',
    'parts'         : '#0f766e',
    'boards'        : '#1d4ed8',
    'board_parts'   : '#2563eb',
    'models'        : '#6d28d9',
    'model_boards'  : '#7c3aed',
    'board_files'   : '#0e7490',
    'quotes'        : '#b45309',
    'quote_items'   : '#92400e',
    'crm_companies' : '#be185d',
    'crm_contacts'  : '#9d174d',
    'crm_deals'     : '#831843',
    'crm_activities': '#701a75',
    'default'       : '#1e293b',
  };

  for (const sheetName in sheetsConfig) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    if (sheet.getLastRow() === 0) {
      const headers = sheetsConfig[sheetName];
      const r = sheet.getRange(1, 1, 1, headers.length);
      r.setValues([headers]);
      r.setFontWeight('bold');
      r.setBackground(headerColors[sheetName] || headerColors['default']);
      r.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return 'v3 データベースの初期化・構築が完了しました！';
}

// ==========================================
// ① OCR / PDF / 写真 → ナレッジ化
// ==========================================

/**
 * Drive上のファイルをOCRしてsourcesシートに格納、AI要約・自動タグ付けも実施
 */
function syncDriveKnowledge() {
  const folderId = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_FOLDER_ID);
  if (!folderId) return 'エラー: KNOWLEDGE_FOLDER_ID が未設定です。';
  const folder = DriveApp.getFolderById(folderId);
  const existingIds = getSheetData('sources').map(r => String(r['ID']));
  let addedCount = 0;

  const processFile = (file) => {
    const fileId = file.getId();
    if (existingIds.includes(fileId)) return;
    const mime = file.getMimeType();
    let text = '', typeName = '';

    if (mime === MimeType.GOOGLE_DOCS) {
      text = DocumentApp.openById(fileId).getBody().getText();
      typeName = 'Google Docs';
    } else if (mime === MimeType.PLAIN_TEXT || mime === MimeType.CSV) {
      text = file.getBlob().getDataAsString('utf-8');
      typeName = 'Text/CSV';
    } else if ([MimeType.PDF, MimeType.JPEG, MimeType.PNG, 'image/tiff', 'image/bmp'].includes(mime)) {
      try {
        const tempDoc = Drive.Files.insert(
          { title: file.getName() + '_OCR_Temp', mimeType: MimeType.GOOGLE_DOCS },
          file.getBlob(), { ocr: true, ocrLanguage: 'ja' }
        );
        text = DocumentApp.openById(tempDoc.id).getBody().getText();
        Drive.Files.remove(tempDoc.id);
        typeName = 'PDF/Image (OCR)';
      } catch(e) {
        text = '[OCRエラー] ' + e.message;
        typeName = 'OCR_Failed';
      }
    } else { return; }

    if (text.length > 40000) text = text.substring(0, 40000) + '\n...(省略)';

    // Gemini で AI 要約・タグ・カテゴリ自動生成
    const meta = _aiSummarizeSource(file.getName(), text);

    appendToSheet('sources', [
      fileId, file.getName(), file.getUrl(), typeName, getCurrentTime(),
      text, meta.summary, meta.category, meta.tags
    ]);
    existingIds.push(fileId);
    addedCount++;
  };

  // 再帰的にサブフォルダもスキャン
  const scanFolder = (f) => {
    const files = f.getFiles();
    while (files.hasNext()) processFile(files.next());
    const subs = f.getFolders();
    while (subs.hasNext()) scanFolder(subs.next());
  };
  scanFolder(folder);

  return `同期完了: ${addedCount} 件を変換・ナレッジ化しました。`;
}

/**
 * 単一ファイルIDを指定してOCR→ナレッジ化（フロントエンドのドラッグ&ドロップ対応）
 */
function importPDFAndClassify(fileId, context) {
  const file = DriveApp.getFileById(fileId);
  const mime = file.getMimeType();
  let text = '', typeName = '';

  if ([MimeType.PDF, MimeType.JPEG, MimeType.PNG, 'image/tiff'].includes(mime)) {
    const tempDoc = Drive.Files.insert(
      { title: file.getName() + '_OCR_Temp', mimeType: MimeType.GOOGLE_DOCS },
      file.getBlob(), { ocr: true, ocrLanguage: 'ja' }
    );
    text = DocumentApp.openById(tempDoc.id).getBody().getText();
    Drive.Files.remove(tempDoc.id);
    typeName = 'PDF/Image (OCR)';
  } else if (mime === MimeType.GOOGLE_DOCS) {
    text = DocumentApp.openById(fileId).getBody().getText();
    typeName = 'Google Docs';
  } else {
    text = file.getBlob().getDataAsString('utf-8');
    typeName = 'Text';
  }

  const meta = _aiSummarizeSource(file.getName(), text, context);
  appendToSheet('sources', [
    fileId, file.getName(), file.getUrl(), typeName, getCurrentTime(),
    text, meta.summary, meta.category, meta.tags
  ]);

  // BOM/見積関連ならBOMシートへ自動取り込みも試みる
  if (meta.category === 'BOM' || meta.category === '見積書') {
    _tryAutoImportBOM(text, meta.category, context);
  }

  return { message: 'OCR・ナレッジ化完了', summary: meta.summary, category: meta.category, tags: meta.tags };
}

/**
 * Gemini でファイルの要約・カテゴリ・タグを自動生成
 */
function _aiSummarizeSource(fileName, text, context) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  if (!apiKey) return { summary: '(APIキー未設定)', category: 'その他', tags: '' };

  const prompt = `以下のドキュメントを分析してください。
ファイル名: ${fileName}
${context ? '追加コンテキスト: ' + context : ''}
テキスト（先頭3000字）:
${text.substring(0, 3000)}

以下のJSON形式のみで返答してください（余分な文字不要）:
{"summary":"100字以内の要約","category":"BOM|見積書|仕様書|議事録|カタログ|図面|その他 のどれか1つ","tags":"カンマ区切りのキーワード5個"}`;

  try {
    const res = _callGemini(prompt, 0.1);
    const cleaned = res.replace(/```json|```/g, '').trim();
    return JSON.parse(cleaned);
  } catch(e) {
    return { summary: text.substring(0, 100), category: 'その他', tags: '' };
  }
}

/**
 * OCRテキストからBOM/見積データを自動抽出してシートに書き込む試み
 */
function _tryAutoImportBOM(text, category, context) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  if (!apiKey) return;

  const prompt = `以下のOCRテキストから部品/製品情報を抽出してください。
カテゴリ: ${category}
テキスト:
${text.substring(0, 5000)}

以下のJSON形式のみで返答してください:
{"items":[{"name":"品名","partNo":"品番","qty":1,"unitPrice":0,"unit":"個","maker":"","spec":""}]}
抽出できない場合は {"items":[]} を返してください。`;

  try {
    const res = _callGemini(prompt, 0.1);
    const cleaned = res.replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(cleaned);
    if (!parsed.items || parsed.items.length === 0) return;

    parsed.items.forEach(item => {
      if (!item.name) return;
      appendToSheet('parts', [
        generateId(), item.name, item.partNo || '', item.maker || '',
        item.spec || '', item.unitPrice || 0, 'JPY', '', '', '[OCR自動取込] ' + (context || ''),
        getCurrentTime()
      ]);
    });
  } catch(e) { /* 失敗しても継続 */ }
}

// ==========================================
// ② BOM マスタ管理
// ==========================================

// --- 基板マスタ ---
function saveBoard(data) {
  const existing = getSheetData('boards').find(r => String(r['基板ID']) === String(data.boardId));
  if (existing) {
    updateSheetRow('boards', '基板ID', data.boardId, [
      data.boardId, data.name, data.revision, data.description, data.parentModelId || '',
      data.cost || 0, data.hours || 0, data.status || '設計中', existing['登録日時']
    ]);
    return '基板マスタを更新しました。';
  }
  appendToSheet('boards', [
    data.boardId || generateId(), data.name, data.revision || 'Rev.1',
    data.description || '', data.parentModelId || '', data.cost || 0,
    data.hours || 0, data.status || '設計中', getCurrentTime()
  ]);
  return '基板マスタに登録しました。';
}

// --- 部品マスタ ---
function savePart(data) {
  const id = data.partId || generateId();
  appendToSheet('parts', [
    id, data.name, data.partNo || '', data.maker || '',
    data.spec || '', data.unitPrice || 0, data.currency || 'JPY',
    data.leadTime || 0, data.stock || 0, data.note || '', getCurrentTime()
  ]);
  return '部品マスタに登録しました (ID: ' + id + ')';
}

// --- 基板-部品 紐付け ---
function saveBoardPart(data) {
  appendToSheet('board_parts', [
    generateId(), data.boardId, data.partId, data.qty || 1,
    data.mountPos || '', data.note || ''
  ]);
  return '基板-部品を紐付けました。';
}

// --- 機種（製品）マスタ ---
function saveModel(data) {
  const id = data.modelId || generateId();
  appendToSheet('models', [
    id, data.name, data.partNo || '', data.category || '',
    data.description || '', data.price || 0, data.cost || 0,
    data.margin || 0, data.status || '現行品', data.successorId || '', getCurrentTime()
  ]);
  return '機種マスタに登録しました (ID: ' + id + ')';
}

// --- 機種-基板 紐付け ---
function saveModelBoard(data) {
  appendToSheet('model_boards', [
    generateId(), data.modelId, data.boardId, data.qty || 1,
    data.role || '', data.note || ''
  ]);
  return '機種-基板を紐付けました。';
}

// --- BOMツリー取得（機種ID → 基板一覧 → 部品一覧） ---
function getBOMTree(modelId) {
  const models    = getSheetData('models');
  const modelBoards = getSheetData('model_boards');
  const boards    = getSheetData('boards');
  const boardParts  = getSheetData('board_parts');
  const parts     = getSheetData('parts');

  const model = models.find(m => String(m['機種ID']) === String(modelId));
  if (!model) return null;

  const mbLinks = modelBoards.filter(r => String(r['機種ID']) === String(modelId));
  const boardNodes = mbLinks.map(mb => {
    const board = boards.find(b => String(b['基板ID']) === String(mb['基板ID'])) || {};
    const bpLinks = boardParts.filter(r => String(r['基板ID']) === String(mb['基板ID']));
    const partNodes = bpLinks.map(bp => {
      const part = parts.find(p => String(p['部品ID']) === String(bp['部品ID'])) || {};
      return { ...bp, partDetail: part };
    });
    return { ...mb, boardDetail: board, parts: partNodes };
  });

  return { model, boards: boardNodes };
}

// --- 見積書マスタ ---
function saveQuote(data) {
  const id = data.quoteId || generateId();
  const quoteNo = data.quoteNo || ('Q-' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd') + '-' + id.substring(0, 4).toUpperCase());
  appendToSheet('quotes', [
    id, quoteNo, data.clientName, data.clientId || '', data.subject || '',
    data.createdDate || getCurrentDate(), data.expiryDate || '',
    data.totalAmount || 0, data.taxRate || 10, data.totalWithTax || 0,
    data.status || '作成中', data.salesRep || '', data.note || ''
  ]);
  return { message: '見積書を作成しました', quoteId: id, quoteNo };
}

function saveQuoteItem(data) {
  appendToSheet('quote_items', [
    generateId(), data.quoteId, data.modelId || '', data.boardId || '',
    data.partId || '', data.itemName, data.qty || 1, data.unitPrice || 0,
    (data.qty || 1) * (data.unitPrice || 0), data.note || ''
  ]);
  return '明細を追加しました。';
}

// BOMから見積明細を自動生成
function buildQuoteBOM(quoteId) {
  const quoteRow = getSheetData('quotes').find(r => String(r['見積ID']) === String(quoteId));
  if (!quoteRow) return 'エラー: 見積IDが見つかりません。';

  // すでに紐づく機種IDがあれば BOMツリーから展開
  const qitems = getSheetData('quote_items').filter(r => String(r['見積ID']) === String(quoteId));
  let total = 0;
  qitems.forEach(item => {
    if (item['機種ID']) {
      const tree = getBOMTree(item['機種ID']);
      if (!tree) return;
      tree.boards.forEach(b => {
        b.parts.forEach(p => {
          const lineAmt = parseFloat(p['数量'] || 0) * parseFloat(p.partDetail['単価'] || 0) * parseFloat(item['数量'] || 1);
          total += lineAmt;
          appendToSheet('quote_items', [
            generateId(), quoteId, '', String(b['基板ID']), String(p['部品ID']),
            String(p.partDetail['部品名'] || ''), parseFloat(p['数量'] || 0) * parseFloat(item['数量'] || 1),
            parseFloat(p.partDetail['単価'] || 0), lineAmt, '[BOM自動展開]'
          ]);
        });
      });
    }
  });

  // 合計を更新
  updateSheetCellByKey('quotes', '見積ID', quoteId, '合計金額', total);
  const taxRate = parseFloat(quoteRow['税率(%)'] || 10) / 100;
  updateSheetCellByKey('quotes', '見積ID', quoteId, '税込合計', Math.round(total * (1 + taxRate)));

  return `BOMから見積明細を展開しました。合計: ¥${total.toLocaleString()}`;
}

// ==========================================
// ③ CRM 機能
// ==========================================

function saveCRMContact(data) {
  const id = data.contactId || generateId();
  appendToSheet('crm_contacts', [
    id, data.companyId || '', data.name, data.title || '',
    data.email || '', data.phone || '', data.salesRep || '',
    getCurrentTime(), data.note || ''
  ]);
  return '連絡先を登録しました (ID: ' + id + ')';
}

function saveCRMCompany(data) {
  const id = data.companyId || generateId();
  appendToSheet('crm_companies', [
    id, data.name, data.industry || '', data.prefecture || '',
    data.address || '', data.phone || '', data.url || '',
    data.salesRep || '', data.rank || 'C', data.note || '', getCurrentTime()
  ]);
  return '会社を登録しました (ID: ' + id + ')';
}

function saveCRMDeal(data) {
  const id = data.dealId || generateId();
  appendToSheet('crm_deals', [
    id, data.companyId, data.contactId || '', data.subject,
    data.phase || 'アプローチ', data.amount || 0, data.probability || 10,
    data.expectedDate || '', data.salesRep || '', data.quoteId || '',
    getCurrentTime(), data.note || ''
  ]);
  return '案件を登録しました (ID: ' + id + ')';
}

function saveCRMActivity(data) {
  const id = data.activityId || generateId();
  appendToSheet('crm_activities', [
    id, data.dealId || '', data.companyId || '', data.contactId || '',
    data.type || '訪問', data.datetime || getCurrentTime(),
    data.content || '', data.nextAction || '', data.salesRep || ''
  ]);
  // 活動内容をmeetingsにも自動登録（ナレッジ共有）
  if (data.content && data.content.length > 20) {
    appendToSheet('meetings', [
      generateId(), getCurrentTime(),
      data.companyName || data.companyId,
      data.relatedProduct || '',
      '[CRM活動] ' + data.type + ': ' + data.content.substring(0, 100),
      data.content, data.nextAction || '', data.salesRep || ''
    ]);
  }
  return '活動を記録しました (ID: ' + id + ')';
}

// CRM ダッシュボード用データ
function getCRMDashboard() {
  const deals    = getSheetData('crm_deals');
  const activities = getSheetData('crm_activities');
  const companies = getSheetData('crm_companies');
  const contacts  = getSheetData('crm_contacts');
  const quotes   = getSheetData('quotes');

  // フェーズ別件数・金額
  const phaseOrder = ['アプローチ', '提案', '見積提示', '交渉', '受注', '失注'];
  const phaseStats = {};
  phaseOrder.forEach(p => { phaseStats[p] = { count: 0, amount: 0 }; });
  deals.forEach(d => {
    const ph = d['フェーズ'] || 'アプローチ';
    if (!phaseStats[ph]) phaseStats[ph] = { count: 0, amount: 0 };
    phaseStats[ph].count++;
    phaseStats[ph].amount += parseFloat(d['金額'] || 0);
  });

  // 最近の活動（30件）
  const recentActivities = activities.slice(-30).reverse();

  // 担当別案件数
  const salesStats = {};
  deals.forEach(d => {
    const rep = d['担当営業'] || '未設定';
    if (!salesStats[rep]) salesStats[rep] = { deals: 0, amount: 0 };
    salesStats[rep].deals++;
    salesStats[rep].amount += parseFloat(d['金額'] || 0);
  });

  return {
    summary: {
      totalDeals     : deals.length,
      totalCompanies : companies.length,
      totalContacts  : contacts.length,
      totalQuotes    : quotes.length,
      pipelineAmount : deals.filter(d => !['受注','失注'].includes(d['フェーズ'])).reduce((s, d) => s + parseFloat(d['金額'] || 0), 0),
      wonAmount      : deals.filter(d => d['フェーズ'] === '受注').reduce((s, d) => s + parseFloat(d['金額'] || 0), 0),
    },
    phaseStats,
    recentActivities,
    salesStats,
    deals: deals.slice(-50).reverse(),
    companies: companies.slice(-50).reverse(),
  };
}

// ==========================================
// BOM ダッシュボード用データ
// ==========================================

function getBOMDashboard() {
  return {
    models     : getSheetData('models'),
    boards     : getSheetData('boards'),
    parts      : getSheetData('parts'),
    boardParts : getSheetData('board_parts'),
    modelBoards: getSheetData('model_boards'),
    quotes     : getSheetData('quotes'),
    quoteItems : getSheetData('quote_items'),
  };
}

// ==========================================
// ナレッジ / AI 検索
// ==========================================

function processAIQuery(query, mode) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  if (!apiKey) return 'APIキーが未設定です。';

  const products = getSheetData('models', 100);
  const boards   = getSheetData('boards', 50);
  const parts    = getSheetData('parts', 100);
  const meetings = getSheetData('meetings', 30);
  const notes    = getSheetData('notes', 30);
  const sources  = getSheetData('sources', 20);
  const deals    = getSheetData('crm_deals', 30);

  let ctx = '【機種マスタ】\n' + products.map(p => `- ${p['機種名']}(${p['型番']}) 価格:${p['販売価格']} 状態:${p['ステータス']}`).join('\n');
  ctx += '\n\n【基板マスタ】\n' + boards.map(b => `- ${b['基板名']} Rev:${b['リビジョン']} コスト:${b['製造コスト']}`).join('\n');
  ctx += '\n\n【商談議事録】\n' + meetings.map(m => `- [${m['登録日時']}] ${m['顧客名']}: ${m['内容サマリ']} 次:${m['次回アクション']}`).join('\n');
  ctx += '\n\n【CRM案件】\n' + deals.map(d => `- ${d['件名']} フェーズ:${d['フェーズ']} 金額:${d['金額']} 確度:${d['確度(%)']}`).join('\n');
  ctx += '\n\n【ナレッジメモ】\n' + notes.map(n => `- [${n['タグ']}] ${n['メモ内容']}`).join('\n');
  ctx += '\n\n【ドキュメント（OCR含む）】\n' + sources.map(s => `- ${s['ファイル名']} [${s['カテゴリ']}] 要約:${s['AI要約']}\n本文抜粋:${String(s['OCRテキスト'] || '').substring(0, 3000)}`).join('\n\n');

  const systemMap = {
    search  : `優秀なセールスAIです。以下の社内データのみを情報源として回答してください。\n${ctx}`,
    roleplay: `厳格な顧客（購買部長）として、営業担当者の提案に鋭く反応してください。背景:\n${ctx}`,
    clone   : `私のデジタルクローンAIです。以下のデータを自分の経験として扱い、私の代わりに営業判断・文章を作成してください。\n${ctx}`,
    analyze_bom: `BOM・見積専門のAIです。部品マスタ・基板・機種情報を踏まえて回答してください。\n${ctx}`,
    analyze_crm: `CRM専門のAIです。案件・活動・顧客情報を踏まえて分析・アドバイスしてください。\n${ctx}`,
  };

  const systemInstruction = systemMap[mode] || systemMap['search'];
  const response = _callGeminiWithSystem(query, systemInstruction, mode === 'roleplay' ? 0.8 : 0.2);
  appendToSheet('qalogs', [getCurrentTime(), query, response]);
  return response;
}

// ==========================================
// ナレッジグラフ
// ==========================================

function getGraphData() {
  const models   = getSheetData('models');
  const meetings = getSheetData('meetings');
  const notes    = getSheetData('notes');
  const sources  = getSheetData('sources');
  const boards   = getSheetData('boards');
  const deals    = getSheetData('crm_deals');
  const edges    = getSheetData('edges');
  const nodes    = [];

  models.forEach(p => nodes.push({ id: String(p['機種ID']), label: String(p['機種名'] || '').substring(0,20), type: 'model', group: 'model', data: p }));
  boards.forEach(b => nodes.push({ id: String(b['基板ID']), label: String(b['基板名'] || '').substring(0,20), type: 'board', group: 'board', data: b }));
  meetings.forEach(m => nodes.push({ id: String(m['ID']), label: String(m['顧客名'] || '').substring(0,20), type: 'meeting', group: 'meeting', data: m }));
  notes.forEach(n => nodes.push({ id: String(n['ID']), label: String(n['メモ内容'] || '').substring(0,20), type: 'note', group: 'note', data: n }));
  sources.forEach(s => nodes.push({ id: String(s['ID']), label: String(s['ファイル名'] || '').substring(0,20), type: 'source', group: 'source', data: s }));
  deals.forEach(d => nodes.push({ id: String(d['案件ID']), label: String(d['件名'] || '').substring(0,20), type: 'deal', group: 'deal', data: d }));

  // model_boards から自動エッジ
  const modelBoards = getSheetData('model_boards');
  const autoEdges = modelBoards.map(mb => ({ id: 'mb_'+mb['ID'], from: String(mb['機種ID']), to: String(mb['基板ID']), label: '搭載', arrows: 'to' }));

  const manualEdges = edges.map(e => ({ id: String(e['ID']), from: String(e['Source_ID']), to: String(e['Target_ID']), label: String(e['Relation_Type'] || ''), arrows: 'to' }));

  return { nodes, edges: [...autoEdges, ...manualEdges] };
}

function saveEdge(data) {
  appendToSheet('edges', [generateId(), data.sourceId, data.targetId, data.relationType || '関連', getCurrentTime()]);
  return 'エッジを登録しました。';
}

function autoExtractEdges() {
  const notes = getSheetData('notes');
  const models = getSheetData('models');
  const edges = getSheetData('edges');
  const existingPairs = new Set(edges.map(e => `${e['Source_ID']}__${e['Target_ID']}`));
  const productNameMap = {};
  models.forEach(p => { if (p['機種名']) productNameMap[String(p['機種名']).toLowerCase()] = String(p['機種ID']); });
  let addedCount = 0;

  notes.forEach(note => {
    const content = String(note['メモ内容'] || '');
    const noteId = String(note['ID']);
    const wikiLinks = content.match(/\[\[([^\]]+)\]\]/g) || [];
    wikiLinks.forEach(link => {
      const linkText = link.replace(/\[\[|\]\]/g, '').toLowerCase().trim();
      const targetId = productNameMap[linkText];
      if (targetId && !existingPairs.has(`${noteId}__${targetId}`)) {
        appendToSheet('edges', [generateId(), noteId, targetId, '[[WikiLink]]', getCurrentTime()]);
        existingPairs.add(`${noteId}__${targetId}`);
        addedCount++;
      }
    });
  });
  return `自動エッジ抽出: ${addedCount}件の関係性を登録しました。`;
}

// ==========================================
// ダッシュボード / テーブル API
// ==========================================

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getCount = (n) => { const s = ss.getSheetByName(n); return s ? Math.max(0, s.getLastRow() - 1) : 0; };
  return {
    sourcesCount  : getCount('sources'),
    meetingsCount : getCount('meetings'),
    notesCount    : getCount('notes'),
    modelsCount   : getCount('models'),
    boardsCount   : getCount('boards'),
    partsCount    : getCount('parts'),
    quotesCount   : getCount('quotes'),
    dealsCount    : getCount('crm_deals'),
    companiesCount: getCount('crm_companies'),
  };
}

function getAllDataForTables() {
  const sheets = ['models','boards','parts','quotes','crm_deals','crm_companies','crm_contacts','meetings','notes','sources'];
  const result = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  sheets.forEach(name => {
    const s = ss.getSheetByName(name);
    if (!s) { result[name] = { headers: [], rows: [] }; return; }
    const data = s.getDataRange().getValues();
    if (data.length <= 1) { result[name] = { headers: data[0] || [], rows: [] }; return; }
    result[name] = { headers: data[0], rows: data.slice(1).map(r => r.map(c => String(c))) };
  });
  return result;
}

// ==========================================
// Google Workspace 統合
// ==========================================

function analyzeWorkspaceAndPushChat() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  if (!apiKey) return 'APIキー未設定';
  const now = new Date();
  const past = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const future = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
  let calCtx = '【カレンダー】\n';
  CalendarApp.getDefaultCalendar().getEvents(past, future).forEach(e => {
    calCtx += `- ${Utilities.formatDate(e.getStartTime(),'JST','MM/dd')} ${e.getTitle()}\n`;
  });
  let mailCtx = '【メール要約】\n';
  GmailApp.search('in:sent OR label:inbox', 0, 20).forEach(t => {
    const m = t.getMessages()[0];
    mailCtx += `- ${m.getSubject()} (${Utilities.formatDate(m.getDate(),'JST','MM/dd')})\n`;
  });
  const prompt = `${calCtx}\n${mailCtx}\n\n上記から: 1)業務傾向 2)顧客トレンド 3)推奨ネクストアクション をマークダウンで報告。`;
  const report = _callGemini(prompt, 0.3);
  const webhookUrl = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_CHAT_WEBHOOK);
  if (webhookUrl) UrlFetchApp.fetch(webhookUrl, { method:'post', contentType:'application/json', payload: JSON.stringify({ text: '📊 自動業務分析レポート\n' + report }), muteHttpExceptions: true });
  return report;
}

// ==========================================
// DB ユーティリティ
// ==========================================

function generateId()     { return Utilities.getUuid(); }
function getCurrentTime() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'); }
function getCurrentDate() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd'); }

function appendToSheet(sheetName, rowData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet) sheet.appendRow(rowData);
}

function getSheetData(sheetName, limit = 0) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  let rows = data.slice(1);
  if (limit > 0) rows = rows.slice(-limit);
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function updateRowData(sheetName, id, headerName, newValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return 'シートが存在しません';
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.indexOf(headerName);
  if (colIndex === -1) return '列が見つかりません';
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, colIndex + 1).setValue(newValue);
      return '更新しました';
    }
  }
  return 'IDが見つかりません';
}

function updateSheetRow(sheetName, keyCol, keyVal, newRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyColIdx = headers.indexOf(keyCol);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyColIdx]) === String(keyVal)) {
      sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
      return;
    }
  }
}

function updateSheetCellByKey(sheetName, keyCol, keyVal, targetCol, newValue) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIdx = headers.indexOf(keyCol);
  const targetIdx = headers.indexOf(targetCol);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIdx]) === String(keyVal)) {
      sheet.getRange(i + 1, targetIdx + 1).setValue(newValue);
      return;
    }
  }
}

function deleteRowData(sheetName, id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return 'シートが存在しません';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return '削除しました';
    }
  }
  return 'IDが見つかりません';
}

// ==========================================
// Gemini API ヘルパー
// ==========================================

function _callGemini(prompt, temperature = 0.2) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature }
  };
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  if (json.candidates && json.candidates.length > 0) return json.candidates[0].content.parts[0].text;
  throw new Error('Gemini APIエラー: ' + JSON.stringify(json));
}

function _callGeminiWithSystem(userPrompt, systemInstruction, temperature = 0.2) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    system_instruction: { parts: { text: systemInstruction } },
    contents: [{ parts: [{ text: userPrompt }] }],
    generationConfig: { temperature }
  };
  const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());
  if (json.candidates && json.candidates.length > 0) return json.candidates[0].content.parts[0].text;
  throw new Error('Gemini APIエラー: ' + JSON.stringify(json));
}

// ==========================================
// Drive ファイルスキャン & 基板ファイルリンク
// ==========================================

function scanAndLinkBoardFiles(folderId, regexStr) {
  if (!folderId) return 'フォルダIDを指定してください。';
  const pattern = regexStr && regexStr.trim() ? new RegExp(regexStr) : /([A-Z0-9\-]+)_(部品表|構成表|BOM|図面|仕様書)/;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('board_files');
  if (!sheet) return 'board_filesシートが存在しません。';
  const existing = {};
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const bid = String(data[i][0]);
    if (bid) existing[bid] = i + 1;
  }
  const colMap = { '部品表': 2, 'BOM': 3, '構成表': 4, '図面': 5, '仕様書': 5 };
  let folder;
  try { folder = DriveApp.getFolderById(folderId); } catch(e) { return 'フォルダID無効: ' + e.message; }
  const now = getCurrentTime();
  let scanned = 0, matched = 0;
  const scanFolder = (f) => {
    const iter = f.getFiles();
    while (iter.hasNext()) {
      const file = iter.next();
      const name = file.getName();
      const url = file.getUrl();
      scanned++;
      const m = name.match(pattern);
      if (!m) return;
      const boardId = m[1]; const fileType = m[2] || 'その他';
      const col = colMap[fileType] || 5;
      matched++;
      if (existing[boardId]) {
        sheet.getRange(existing[boardId], col).setValue(url);
        sheet.getRange(existing[boardId], 6).setValue(now);
      } else {
        const newRow = [boardId, '', '', '', '', now];
        newRow[col - 1] = url;
        sheet.appendRow(newRow);
        existing[boardId] = sheet.getLastRow();
      }
    }
    const subs = f.getFolders();
    while (subs.hasNext()) scanFolder(subs.next());
  };
  scanFolder(folder);
  return `スキャン完了: ${scanned}件中 ${matched}件がマッチしました。`;
}

// ==========================================
// キャリア目標
// ==========================================

function saveGoal(data) {
  appendToSheet('goals', [generateId(), getCurrentTime(), data.period, data.title, data.kpi, data.progress || 0, data.eval || '', data.refId || '']);
  return 'キャリア目標を登録しました。';
}

function getCareerData() {
  return { goals: getSheetData('goals'), stats: { meetings: getSheetData('meetings').length, sources: getSheetData('sources').length } };
}

// ==========================================
// インポート (CSV / 一括)
// ==========================================

function importBulkData(sheetType, headers, dataRows) {
  const SHEET_CONFIG = {
    models : {
      sheetName: 'models', required: ['機種名'],
      build: (h, r) => [generateId(), getCellValue(h,r,'機種名'), getCellValue(h,r,'型番'), getCellValue(h,r,'カテゴリ'), getCellValue(h,r,'説明'), getCellValue(h,r,'販売価格')||0, getCellValue(h,r,'原価')||0, 0, getCellValue(h,r,'ステータス')||'現行品', '', getCurrentTime()]
    },
    boards : {
      sheetName: 'boards', required: ['基板名'],
      build: (h, r) => [generateId(), getCellValue(h,r,'基板名'), getCellValue(h,r,'リビジョン')||'Rev.1', getCellValue(h,r,'説明'), '', getCellValue(h,r,'製造コスト')||0, 0, getCellValue(h,r,'ステータス')||'設計中', getCurrentTime()]
    },
    parts : {
      sheetName: 'parts', required: ['部品名'],
      build: (h, r) => [generateId(), getCellValue(h,r,'部品名'), getCellValue(h,r,'部品番号'), getCellValue(h,r,'メーカー'), getCellValue(h,r,'仕様'), getCellValue(h,r,'単価')||0, 'JPY', getCellValue(h,r,'リードタイム(日)')||0, getCellValue(h,r,'在庫数')||0, '', getCurrentTime()]
    },
    crm_companies : {
      sheetName: 'crm_companies', required: ['会社名'],
      build: (h, r) => [generateId(), getCellValue(h,r,'会社名'), getCellValue(h,r,'業種'), getCellValue(h,r,'都道府県'), getCellValue(h,r,'住所'), getCellValue(h,r,'電話'), getCellValue(h,r,'URL'), getCellValue(h,r,'担当営業'), getCellValue(h,r,'ランク')||'C', '', getCurrentTime()]
    },
    quotes : {
      sheetName: 'quotes', required: ['顧客名'],
      build: (h, r) => [generateId(), getCellValue(h,r,'見積番号'), getCellValue(h,r,'顧客名'), '', getCellValue(h,r,'件名'), getCellValue(h,r,'作成日')||getCurrentDate(), '', getCellValue(h,r,'合計金額')||0, 10, 0, getCellValue(h,r,'ステータス')||'作成中', getCellValue(h,r,'担当営業'), '']
    },
  };
  const config = SHEET_CONFIG[sheetType];
  if (!config) throw new Error('不明な種別: ' + sheetType);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
  if (!sheet) throw new Error('シートが存在しません: ' + config.sheetName);
  let ok = 0, skip = 0;
  for (const row of dataRows) {
    if (row.every(c => !String(c).trim())) { skip++; continue; }
    try { sheet.appendRow(config.build(headers, row)); ok++; } catch(e) { skip++; }
  }
  return `インポート完了: ${ok}件登録, ${skip}件スキップ`;
}

function getCellValue(headers, row, col) {
  const idx = headers.indexOf(col);
  return idx === -1 ? '' : String(row[idx] || '').trim();
}

function readFromSpreadsheet(spreadsheetId, tabName) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = tabName ? ss.getSheetByName(tabName) || ss.getSheets()[0] : ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { headers: [], rows: [] };
    return { headers: data[0].map(h => String(h).trim()), rows: data.slice(1).filter(r => r.some(c => String(c).trim())).map(r => r.map(c => String(c))) };
  } catch(e) {
    throw new Error('スプレッドシートへのアクセス失敗: ' + e.message);
  }
}

// カレンダー表示用イベント取得
function getCalendarEvents() {
  const events = [];
  const extractDate = (str) => {
    if (!str) return '';
    const m = String(str).match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    return m ? `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}` : '';
  };
  getSheetData('meetings').forEach(m => {
    const d = extractDate(m['登録日時']);
    if (d) events.push({ id: m['ID'], type: 'meeting', date: d, client: m['顧客名'], summary: m['内容サマリ'], rep: m['担当営業'] });
  });
  getSheetData('crm_activities').forEach(a => {
    const d = extractDate(a['日時']);
    if (d) events.push({ id: a['活動ID'], type: 'activity', date: d, client: a['会社ID'], summary: String(a['内容']).substring(0,80), rep: a['担当営業'] });
  });
  return events;
}
