// ==========================================
// DriveClassifier.gs
// Drive自動分類・階層管理システム
// 既存のCode.gsに追記して使用
// ==========================================

// ==========================================
// 【DB拡張】初回セットアップに追加するシート定義
// setupDatabase() の sheetsConfig に以下を追加してください:
//
// 'models': ['機種ID', '機種名', 'Drive フォルダID', 'Drive URL', '登録日時', '備考'],
// 'boards': ['基板ID', '親機種ID', '基板名', 'DriveフォルダID', 'Drive URL', '登録日時'],
// 'parts': ['部品ID', '親基板ID', '部品名', '型番', '単価', 'DriveファイルID', 'Drive URL', '登録日時'],
// 'quotes': ['見積ID', '親機種ID', '見積番号', '顧客名', 'DriveファイルID', 'Drive URL', 'ステータス', '登録日時'],
// 'doc_hierarchy': ['ドキュメントID', '親ID', '親種別(model/board)', '種別', 'ファイル名', 'DriveファイルID', 'Drive URL', '登録日時'],
// ==========================================

const CLASSIFIER_RULES = {
  // ファイル名パターン → 種別の自動判定ルール
  model_spec:   { patterns: ['製品仕様', '機種仕様', 'spec', '仕様書'], label: '機種製品仕様書',   icon: '📋' },
  board_design: { patterns: ['基板設計', '回路設計', '基板仕様', 'schematic', 'PCB'], label: '基板設計仕様書', icon: '🔧' },
  bom:          { patterns: ['BOM', '構成表', 'bill_of_material', '部品表', 'parts_list'], label: '構成表(BOM)', icon: '📦' },
  quote:        { patterns: ['見積', 'quote', 'QUOTE', '見積書', 'estimate'], label: '見積書', icon: '💴' },
  drawing:      { patterns: ['図面', 'drawing', 'DWG', 'CAD', '外形図'], label: '図面', icon: '📐' },
  manual:       { patterns: ['マニュアル', 'manual', '操作説明', '取扱説明'], label: 'マニュアル', icon: '📖' },
  test:         { patterns: ['試験', 'テスト', 'test', '検査', '評価'], label: '試験・検査書', icon: '🧪' },
};

// ==========================================
// メイン: Driveフォルダを自動スキャン＆分類
// ==========================================

/**
 * Drive上の指定フォルダを再帰スキャンし、
 * ファイル名から種別を自動判定してスプレッドシートに転記する
 * @param {string} rootFolderId - スキャンするDriveフォルダのID
 * @param {string} modelName    - 紐付ける機種名 (UI側から受け取る)
 */
function scanAndClassifyDriveFolder(rootFolderId, modelName) {
  if (!rootFolderId) return { status: 'error', message: 'フォルダIDが未入力です。' };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureClassifierSheets(ss);

  try {
    const folder = DriveApp.getFolderById(rootFolderId);

    // 1. 機種レコードを生成 or 取得
    const modelId = getOrCreateModel(ss, modelName || folder.getName(), rootFolderId, folder.getUrl());

    // 2. フォルダ内を再帰スキャン
    const results = { classified: [], unclassified: [] };
    scanFolderRecursive(ss, folder, modelId, null, results);

    // 3. 結果サマリーを返す
    return {
      status: 'success',
      modelId: modelId,
      modelName: modelName || folder.getName(),
      classified: results.classified.length,
      unclassified: results.unclassified.length,
      details: results.classified,
      unclassifiedFiles: results.unclassified,
      message: `✅ 分類完了: ${results.classified.length}件 自動分類 / ${results.unclassified.length}件 未分類`
    };
  } catch(e) {
    return { status: 'error', message: 'スキャンエラー: ' + e.message };
  }
}

/**
 * フォルダを再帰的にスキャンし、ファイルを分類する
 */
function scanFolderRecursive(ss, folder, modelId, parentBoardId, results) {
  const files = folder.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const fileId   = file.getId();
    const fileUrl  = file.getUrl();
    const mimeType = file.getMimeType();

    // ファイル種別を自動判定
    const classification = classifyFile(fileName, mimeType);

    if (classification) {
      // 判定できた場合: 種別ごとにシートへ転記
      saveClassifiedFile(ss, file, classification, modelId, parentBoardId);
      results.classified.push({ name: fileName, type: classification.label, icon: classification.icon });
    } else {
      // 判定できなかった場合: 未分類リストへ
      saveUnclassifiedFile(ss, file, modelId);
      results.unclassified.push({ name: fileName, url: fileUrl });
    }
  }

  // サブフォルダも再帰処理
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    const subName = sub.getName();
    
    // フォルダ名から「基板」サブフォルダか判定
    let newParentBoardId = parentBoardId;
    if (isBoardFolder(subName)) {
      newParentBoardId = getOrCreateBoard(ss, subName, modelId, sub.getId(), sub.getUrl());
    }
    
    scanFolderRecursive(ss, sub, modelId, newParentBoardId, results);
  }
}

// ==========================================
// 自動分類ロジック
// ==========================================

/**
 * ファイル名・MIMEタイプからドキュメント種別を判定する
 */
function classifyFile(fileName, mimeType) {
  const lowerName = fileName.toLowerCase().replace(/[-_\s]/g, '');

  for (const [key, rule] of Object.entries(CLASSIFIER_RULES)) {
    for (const pattern of rule.patterns) {
      if (lowerName.includes(pattern.toLowerCase().replace(/[-_\s]/g, ''))) {
        return { key: key, label: rule.label, icon: rule.icon };
      }
    }
  }

  // MIMEタイプでの補助判定
  if (mimeType === MimeType.GOOGLE_SHEETS || fileName.match(/\.xlsx?$/i)) {
    return CLASSIFIER_RULES.bom; // スプレッドシートはデフォルトでBOM扱い
  }
  
  return null; // 判定不能
}

/**
 * フォルダ名が「基板」関連かどうか判定
 */
function isBoardFolder(folderName) {
  const keywords = ['基板', 'board', 'PCB', 'substrate', 'ボード'];
  return keywords.some(k => folderName.toLowerCase().includes(k.toLowerCase()));
}

// ==========================================
// シートへの書き込み処理
// ==========================================

function getOrCreateModel(ss, modelName, folderId, folderUrl) {
  const sheet = ss.getSheetByName('models');
  const data = sheet.getDataRange().getValues();
  
  // 既存の機種IDを検索 (フォルダIDで一意判定)
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === folderId) return data[i][0];
  }
  
  // 新規作成
  const newId = generateId();
  sheet.appendRow([newId, modelName, folderId, folderUrl, getCurrentTime(), '']);
  return newId;
}

function getOrCreateBoard(ss, boardName, modelId, folderId, folderUrl) {
  const sheet = ss.getSheetByName('boards');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === folderId) return data[i][0];
  }
  
  const newId = generateId();
  sheet.appendRow([newId, modelId, boardName, folderId, folderUrl, getCurrentTime()]);
  return newId;
}

function saveClassifiedFile(ss, file, classification, modelId, boardId) {
  const sheet = ss.getSheetByName('doc_hierarchy');
  
  // 重複チェック (DriveファイルIDで判定)
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === file.getId()) return; // 既登録ならスキップ
  }

  // 見積書の場合は quotesシートにも転記
  if (classification.key === 'quote') {
    saveToQuoteSheet(ss, file, modelId);
  }

  sheet.appendRow([
    generateId(),
    boardId || modelId,
    boardId ? 'board' : 'model',
    classification.label,
    file.getName(),
    file.getId(),
    file.getUrl(),
    getCurrentTime()
  ]);
}

function saveToQuoteSheet(ss, file, modelId) {
  const sheet = ss.getSheetByName('quotes');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === file.getId()) return;
  }
  // 見積番号をファイル名から抽出 (例: "見積書_QUOTE-202.pdf" → "QUOTE-202")
  const quoteNumMatch = file.getName().match(/([A-Z]+-?\d+)/);
  const quoteNum = quoteNumMatch ? quoteNumMatch[1] : file.getName();
  sheet.appendRow([generateId(), modelId, quoteNum, '', file.getId(), file.getUrl(), '未確認', getCurrentTime()]);
}

function saveUnclassifiedFile(ss, file, modelId) {
  const sheet = ss.getSheetByName('doc_hierarchy');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === file.getId()) return;
  }
  sheet.appendRow([generateId(), modelId, 'model', '【未分類】', file.getName(), file.getId(), file.getUrl(), getCurrentTime()]);
}

// ==========================================
// シートの自動生成（初回セットアップ補完）
// ==========================================

function ensureClassifierSheets(ss) {
  const required = {
    'models':        ['機種ID', '機種名', 'DriveフォルダID', 'Drive URL', '登録日時', '備考'],
    'boards':        ['基板ID', '親機種ID', '基板名', 'DriveフォルダID', 'Drive URL', '登録日時'],
    'parts':         ['部品ID', '親基板ID', '部品名', '型番', '単価', 'DriveファイルID', 'Drive URL', '登録日時'],
    'quotes':        ['見積ID', '親機種ID', '見積番号', '顧客名', 'DriveファイルID', 'Drive URL', 'ステータス', '登録日時'],
    'doc_hierarchy': ['ドキュメントID', '親ID', '親種別', 'ファイル種別', 'ファイル名', 'DriveファイルID', 'Drive URL', '登録日時'],
  };

  for (const [name, headers] of Object.entries(required)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      const range = sheet.getRange(1, 1, 1, headers.length);
      range.setValues([headers]);
      range.setFontWeight('bold');
      range.setBackground('#0f172a');
      range.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
}

// ==========================================
// UI向けデータ取得 API
// ==========================================

/**
 * 機種ツリー全体を返す（UI側でツリー表示に使用）
 */
function getModelTree() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureClassifierSheets(ss);

  const models   = getSheetData('models');
  const boards   = getSheetData('boards');
  const docs     = getSheetData('doc_hierarchy');
  const quotes   = getSheetData('quotes');

  return models.map(model => {
    const modelBoards = boards.filter(b => b['親機種ID'] === model['機種ID']);
    const modelDocs   = docs.filter(d => d['親ID'] === model['機種ID'] && d['親種別'] === 'model');
    const modelQuotes = quotes.filter(q => q['親機種ID'] === model['機種ID']);

    return {
      id:      model['機種ID'],
      name:    model['機種名'],
      url:     model['Drive URL'],
      folderId: model['DriveフォルダID'],
      docs:    modelDocs,
      quotes:  modelQuotes,
      boards:  modelBoards.map(board => {
        const boardDocs = docs.filter(d => d['親ID'] === board['基板ID'] && d['親種別'] === 'board');
        return {
          id:   board['基板ID'],
          name: board['基板名'],
          url:  board['Drive URL'],
          docs: boardDocs
        };
      })
    };
  });
}

/**
 * 手動で種別を再分類する（未分類ファイルの修正用）
 */
function reclassifyDocument(docId, newType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return updateRowData('doc_hierarchy', docId, 'ファイル種別', newType);
}

/**
 * DriveフォルダIDの検証（UI入力値チェック用）
 */
function validateDriveFolderId(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    return { valid: true, name: folder.getName(), url: folder.getUrl() };
  } catch(e) {
    return { valid: false, message: 'フォルダが見つかりません。IDを確認してください。' };
  }
}
