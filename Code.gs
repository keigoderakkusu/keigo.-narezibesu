// ==========================================
// Code.gs v2.0 — ナレッジグラフ対応版
// バックエンドロジック・APIルーティング
// ==========================================

const SCRIPT_PROP_GEMINI_KEY = 'GEMINI_API_KEY';
const SCRIPT_PROP_FOLDER_ID = 'KNOWLEDGE_FOLDER_ID';
const SCRIPT_PROP_CHAT_WEBHOOK = 'CHAT_WEBHOOK_URL';

/**
 * Webアプリへのアクセス時の処理
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('ナレッジグラフ Hub v2')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スマホ連携・n8nからの外部WebHookリクエスト受け口 (API機能)
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'saveMeeting') {
      const result = saveMeeting(data.payload);
      return ContentService.createTextOutput(JSON.stringify({status: 'success', message: result})).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'processQuery') {
      const result = processAIQuery(data.payload.query, data.payload.mode || 'clone');
      return ContentService.createTextOutput(JSON.stringify({status: 'success', response: result})).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'saveModel')   { return _jsonOut(saveModel(data.payload));
    } else if (action === 'saveBoard')   { return _jsonOut(saveBoard(data.payload));
    } else if (action === 'saveConfig')  { return _jsonOut(saveConfig(data.payload));
    } else if (action === 'savePart')    { return _jsonOut(savePart(data.payload));
    } else if (action === 'deleteModel') { return _jsonOut({message: deleteModel(data.payload.id)});
    } else if (action === 'deleteBoard') { return _jsonOut({message: deleteBoard(data.payload.id)});
    } else if (action === 'importCSV')   { return _jsonOut(importCSV(data.payload.csvText, data.payload.importType, data.payload.context));
    } else if (action === 'importPDF')   { return _jsonOut(importPDFtoBOM(data.payload.fileId, data.payload.context));
    } else {
      return ContentService.createTextOutput(JSON.stringify({status: 'error', message: '不明なアクション: ' + action})).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 別ファイル(CSS/JavaScript)を動的に読み込むための関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function _jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 【DB自動構築ロジック】初回実行時のみ実行
 * v2: edgesシートを追加
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetsConfig = {
    'products':    ['ID', '製品名', '価格', 'ステータス', '改廃理由', '後継機種', '競合優位性'],
    'meetings':    ['ID', '登録日時', '顧客名', '関連基板・機種', '内容サマリ', '議事録全文'],
    'notes':       ['ID', '登録日時', 'タグ(Obsidianタグ等)', 'メモ内容'],
    'notebooklm':  ['ID', 'ノートブック名', '関連製品', '共有URL', 'Enterprise連携フラグ'],
    'qalogs':      ['日時', 'ユーザー入力', 'AI回答内容'],
    'sources':     ['ID', 'ファイル名', 'URL', 'タイプ', '連携日時', 'テキスト内容'],
    'goals':       ['ID', '登録日時', '対象期間', '目標タイトル', '測定指標(KPI)', '進捗(%)', '達成状況・自己評価', '関連成果ID(議事録等)'],
    // v2: グラフのエッジ管理シート
    'edges':       ['ID', 'Source_ID', 'Target_ID', 'Relation_Type', '登録日時'],
    // v3 新規追加: 基板ファイルリンクDB
    'board_files': ['基板ID', '部品表URL', '構成表URL', 'BOM_URL', 'その他URL', 'スキャン日時', '生ファイル一覧(JSON)']
  };

  const headerColors = {
    'products':    '#1d4ed8',
    'meetings':    '#b45309',
    'notes':       '#15803d',
    'sources':     '#7c3aed',
    'edges':       '#be123c',
    'board_files': '#0e7490',
    'default':     '#1e293b'
  };

  for (const sheetName in sheetsConfig) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    const headers = sheetsConfig[sheetName];
    
    if (sheet.getLastRow() === 0) {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground(headerColors[sheetName] || headerColors['default']);
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  // 遊技機基板ドキュメント管理シートも同時に初期化
  setupGamingDocDatabase();
  return 'データベースの初期化・構築が完了しました！edges・board_files・documentsシートも追加されました。';
}

// ----------------------------------------------------
// ★ 基板ファイル自動スキャン & リンクDB (v3 新規追加)
// ----------------------------------------------------

/**
 * Driveフォルダをスキャンし、命名規則に従って基板ごとにファイルを紐付ける
 * @param {string} folderId  - スキャン対象のGoogle DriveフォルダID
 * @param {string} regexStr  - キャプチャグループ1=基板ID, グループ2=ファイル種類
 *   デフォルト例: ([A-Z0-9\-]+)_(部品表|構成表|BOM|図面|仕様書)
 */
function scanAndLinkBoardFiles(folderId, regexStr) {
  if (!folderId) return 'エラー: フォルダIDを指定してください。';

  // デフォルト命名規則パターン（ユーザーがUIからカスタマイズ可能）
  const pattern = regexStr && regexStr.trim()
    ? new RegExp(regexStr)
    : /([A-Z0-9\-]+)_(部品表|構成表|BOM|図面|仕様書)/;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('board_files');
  if (!sheet) return 'エラー: board_filesシートが存在しません。setupDatabase()を実行してください。';

  // 既存データをオブジェクト化（基板IDをキーに）
  const existing = {};
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const boardId = String(data[i][0]);
    if (boardId) existing[boardId] = i + 1; // 行番号（1-indexed）
  }

  const colMap = {
    '部品表':  2,  // B列
    'BOM':    3,  // C列
    '構成表':  4,  // D列
    '図面':   5,  // E列
    '仕様書':  5   // E列（その他として共用）
  };

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch(e) {
    return 'エラー: フォルダIDが無効です。詳細: ' + e.message;
  }

  const now = getCurrentTime();
  let scannedCount = 0;
  let matchedCount = 0;
  const unmatchedFiles = [];

  // 再帰的にサブフォルダもスキャンするイテレータ処理
  const fileIterators = [folder.getFiles()];
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    fileIterators.push(subFolders.next().getFiles());
  }

  fileIterators.forEach(iter => {
    while (iter.hasNext()) {
      const file  = iter.next();
      const name  = file.getName();
      const url   = file.getUrl();
      const mime  = file.getMimeType();
      scannedCount++;

      const m = name.match(pattern);
      if (!m) {
        unmatchedFiles.push(name);
        continue;
      }

      const boardId  = m[1];
      const fileType = m[2] || 'その他';
      const targetCol = colMap[fileType] || 5;
      matchedCount++;

      if (existing[boardId]) {
        // 既存行を更新
        const row = existing[boardId];
        sheet.getRange(row, targetCol).setValue(url);
        sheet.getRange(row, 6).setValue(now);
      } else {
        // 新規行を追加
        const newRow = [boardId, '', '', '', '', now, ''];
        newRow[targetCol - 1] = url;
        sheet.appendRow(newRow);
        // 行番号を登録（同一基板の後続ファイルのため）
        existing[boardId] = sheet.getLastRow();
      }

      // グラフのnodesとして products/sources にも自動連携
      // sourcesシートに未登録なら追加（グラフノード化）
      const sourceRows = getSheetData('sources');
      const alreadyInSources = sourceRows.some(r => String(r['URL']) === url);
      if (!alreadyInSources) {
        appendToSheet('sources', [
          generateId(), name, url,
          'Drive(' + fileType + ')',
          now,
          '[基板ID: ' + boardId + '] 自動スキャンによる取り込み'
        ]);
      }
    }
  });

  return `スキャン完了: ${scannedCount}件中 ${matchedCount}件が命名規則に一致。\n未マッチ: ${unmatchedFiles.length}件 (${unmatchedFiles.slice(0,3).join(', ')}${unmatchedFiles.length > 3 ? '...' : ''})\n\n基板ファイルDBを更新しました。グラフに反映するには「再読込」してください。`;
}

/**
 * board_filesシートのデータをUI向けに返す
 */
function getBoardFilesData() {
  const data = getSheetData('board_files');
  return data.map(row => ({
    boardId:  row['基板ID'],
    bomUrl:   row['部品表URL']    || row['BOM_URL'] || '',
    kousei:   row['構成表URL']    || '',
    otherUrl: row['その他URL']    || '',
    scanned:  row['スキャン日時'] || ''
  }));
}

// ----------------------------------------------------
// DBユーティリティ関数
// ----------------------------------------------------

function generateId() {
  return Utilities.getUuid();
}

function getCurrentTime() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

function appendToSheet(sheetName, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.appendRow(rowData);
  }
}

function getSheetData(sheetName, limit = 0) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  let rows = data.slice(1);
  
  if (limit > 0) {
    rows = rows.slice(-limit);
  }
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

// ----------------------------------------------------
// ★ ナレッジグラフ API (v2 新規追加)
// ----------------------------------------------------

/**
 * 全ノード（4種）とエッジをまとめて返す
 * vis-networkで描画するためのJSON
 */
function getGraphData() {
  const products  = getSheetData('products');
  const meetings  = getSheetData('meetings');
  const notes     = getSheetData('notes');
  const sources   = getSheetData('sources');
  const edges     = getSheetData('edges');

  const nodes = [];

  // 製品ノード (青系)
  products.forEach(p => {
    nodes.push({
      id:    String(p['ID']),
      label: String(p['製品名'] || '').substring(0, 20),
      title: `製品: ${p['製品名']}\n価格: ${p['価格']}\n状態: ${p['ステータス']}`,
      type:  'product',
      group: 'product',
      data:  p
    });
  });

  // 議事録ノード (オレンジ系)
  meetings.forEach(m => {
    nodes.push({
      id:    String(m['ID']),
      label: String(m['顧客名'] || '').substring(0, 20),
      title: `議事録: ${m['顧客名']}\n${m['登録日時']}\n${m['内容サマリ']}`,
      type:  'meeting',
      group: 'meeting',
      data:  m
    });
  });

  // メモノード (緑系)
  notes.forEach(n => {
    const labelStr = String(n['メモ内容'] || '').substring(0, 20);
    nodes.push({
      id:    String(n['ID']),
      label: labelStr || 'メモ',
      title: `メモ: ${n['タグ(Obsidianタグ等)']}\n${String(n['メモ内容'] || '').substring(0, 100)}`,
      type:  'note',
      group: 'note',
      data:  n
    });
  });

  // ソースノード (紫系)
  sources.forEach(s => {
    nodes.push({
      id:    String(s['ID']),
      label: String(s['ファイル名'] || '').substring(0, 20),
      title: `ドキュメント: ${s['ファイル名']}\nタイプ: ${s['タイプ']}`,
      type:  'source',
      group: 'source',
      data:  s
    });
  });

  // エッジ変換
  const edgeList = edges.map(e => ({
    id:     String(e['ID']),
    from:   String(e['Source_ID']),
    to:     String(e['Target_ID']),
    label:  String(e['Relation_Type'] || ''),
    arrows: 'to'
  }));

  return { nodes, edges: edgeList };
}

/**
 * ノードIDから詳細データを返す（サイドパネル表示用）
 */
function getNodeDetail(nodeId) {
  const allSheets = ['products', 'meetings', 'notes', 'sources'];
  for (const sheetName of allSheets) {
    const rows = getSheetData(sheetName);
    const found = rows.find(r => String(r['ID']) === String(nodeId));
    if (found) {
      // このノードに繋がるエッジも取得
      const edges = getSheetData('edges');
      const relatedEdges = edges.filter(e =>
        String(e['Source_ID']) === String(nodeId) ||
        String(e['Target_ID']) === String(nodeId)
      );
      return { sheetName, data: found, edges: relatedEdges };
    }
  }
  return null;
}

/**
 * 手動でエッジを追加する
 */
function saveEdge(data) {
  // data: { sourceId, targetId, relationType }
  appendToSheet('edges', [
    generateId(),
    data.sourceId,
    data.targetId,
    data.relationType || '関連',
    getCurrentTime()
  ]);
  return 'エッジ（関係性）を登録しました。';
}

/**
 * メモ内の [[WikiLink]] 形式を解析して自動エッジ抽出
 * また、共通タグを持つノード間に自動エッジを生成
 */
function autoExtractEdges() {
  const notes    = getSheetData('notes');
  const products = getSheetData('products');
  const edges    = getSheetData('edges');
  let addedCount = 0;

  // 既存エッジのペアを収集（重複防止）
  const existingPairs = new Set(
    edges.map(e => `${e['Source_ID']}__${e['Target_ID']}`)
  );

  // 全製品名をインデックス化
  const productNameMap = {};
  products.forEach(p => {
    if (p['製品名']) productNameMap[String(p['製品名']).toLowerCase()] = String(p['ID']);
  });

  notes.forEach(note => {
    const content = String(note['メモ内容'] || '');
    const noteId  = String(note['ID']);

    // [[製品名]] or [[メモタイトル]] 形式の自動リンク
    const wikiLinks = content.match(/\[\[([^\]]+)\]\]/g) || [];
    wikiLinks.forEach(link => {
      const linkText = link.replace(/\[\[|\]\]/g, '').toLowerCase().trim();
      const targetProductId = productNameMap[linkText];
      if (targetProductId && !existingPairs.has(`${noteId}__${targetProductId}`)) {
        appendToSheet('edges', [generateId(), noteId, targetProductId, '[[WikiLink]]', getCurrentTime()]);
        existingPairs.add(`${noteId}__${targetProductId}`);
        addedCount++;
      }
    });

    // タグの共通性からメモ同士を関連付け
    const noteTags = String(note['タグ(Obsidianタグ等)'] || '')
      .split(/[\s,#]+/)
      .filter(t => t.length > 1)
      .map(t => t.toLowerCase());

    notes.forEach(otherNote => {
      if (String(otherNote['ID']) === noteId) return;
      const otherTags = String(otherNote['タグ(Obsidianタグ等)'] || '')
        .split(/[\s,#]+/)
        .filter(t => t.length > 1)
        .map(t => t.toLowerCase());

      const sharedTags = noteTags.filter(t => otherTags.includes(t));
      const otherId = String(otherNote['ID']);
      
      if (sharedTags.length > 0 && !existingPairs.has(`${noteId}__${otherId}`)) {
        appendToSheet('edges', [
          generateId(), noteId, otherId,
          `共通タグ: ${sharedTags[0]}`, getCurrentTime()
        ]);
        existingPairs.add(`${noteId}__${otherId}`);
        addedCount++;
      }
    });
  });

  return `自動エッジ抽出完了: ${addedCount}件の関係性を新たに登録しました。`;
}

// ----------------------------------------------------
// インライン編集・データ一括取得 API (SaaS拡張)
// ----------------------------------------------------

function getAllDataForTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getSheetAll = (name) => {
    const s = ss.getSheetByName(name);
    if (!s) return { headers: [], rows: [] };
    const data = s.getDataRange().getValues();
    if (data.length <= 1) return { name: name, headers: data[0] || [], rows: [] };
    return { name: name, headers: data[0], rows: data.slice(1).map(r => r.map(c => String(c))) };
  };

  return {
    products: getSheetAll('products'),
    meetings: getSheetAll('meetings'),
    notes:    getSheetAll('notes')
  };
}

function updateRowData(sheetName, id, headerName, newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'エラー: シートが存在しません';
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 'エラー: データがありません';
  
  const headers = data[0];
  const colIndex = headers.indexOf(headerName);
  if (colIndex === -1) return 'エラー: 該当する列が見つかりません';
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, colIndex + 1).setValue(newValue);
      return '更新しました';
    }
  }
  return 'エラー: 対象のIDが見つかりません';
}

function deleteRowData(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'エラー: シートが存在しません';
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return '削除しました';
    }
  }
  return 'エラー: 対象のIDが見つかりません';
}

// ----------------------------------------------------
// キャリア目標・人事評価・スキルアップ管理 API
// ----------------------------------------------------

function getCareerData() {
  const goals = getSheetData('goals');
  const meetings = getSheetData('meetings');
  
  const now = new Date();
  const oneMonthAgo = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
  
  const recentActivities = meetings.filter(m => {
    const d = new Date(m['登録日時']);
    return d > oneMonthAgo;
  });

  return {
    goals: goals,
    stats: {
      recentMeetings: recentActivities.length,
      knowledgeCreated: getSheetData('notes').length + getSheetData('sources').length
    }
  };
}

function saveGoal(data) {
  appendToSheet('goals', [
    generateId(), getCurrentTime(), data.period, data.title, data.kpi, data.progress, data.eval, data.refId
  ]);
  return 'キャリア目標を登録しました。';
}

// ----------------------------------------------------
// 通勤学習用 音声化リクエスト API (n8n連携)
// ----------------------------------------------------

function requestAudioConversion(title, text) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('AUDIO_GEN_WEBHOOK_URL');
  if (!webhookUrl) return 'エラー: 音声生成用Webhook URLが設定されていません。';

  const payload = {
    'title': title,
    'text': text.substring(0, 5000),
    'userName': Session.getActiveUser().getEmail()
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    UrlFetchApp.fetch(webhookUrl, options);
    return '音声生成をリクエストしました。数分後、Driveの通勤学習フォルダを確認してください。';
  } catch(e) {
    return 'リクエストエラー: ' + e.message;
  }
}

// ----------------------------------------------------
// フロントエンド用 データ取得系API
// ----------------------------------------------------

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getCount = (name) => {
    const s = ss.getSheetByName(name);
    return s ? Math.max(0, s.getLastRow() - 1) : 0;
  };

  const notebookLMs = getSheetData('notebooklm').reverse();

  return {
    productsCount: getCount('products'),
    meetingsCount: getCount('meetings'),
    notesCount:    getCount('notes'),
    edgesCount:    getCount('edges'),
    notebookLMs:   notebookLMs
  };
}

// ----------------------------------------------------
// フロントエンド用 データ登録系API
// ----------------------------------------------------

function saveProduct(data) {
  appendToSheet('products', [
    generateId(), data.name, data.price, data.status, data.reason, data.successor, data.advantage
  ]);
  return '製品情報を正常に登録しました。';
}

function saveMeeting(data) {
  appendToSheet('meetings', [
    generateId(), getCurrentTime(), data.client, data.related, data.summary, data.fulltext
  ]);
  return '商談議事録を正常に登録しました。AIの学習対象に追加されました。';
}

function saveNote(data) {
  appendToSheet('notes', [
    generateId(), getCurrentTime(), data.tags, data.content
  ]);
  return 'ナレッジメモ（Obsidian形式等）を正常に登録しました。';
}

function saveNotebookLM(data) {
  appendToSheet('notebooklm', [
    generateId(), data.name, data.related, data.url, '未連携(手動)'
  ]);
  return 'NotebookLMリンク台帳に登録しました。';
}

// ----------------------------------------------------
// Drive連携機能 (PDF / Docの自動テキスト化とナレッジ格納)
// ----------------------------------------------------

function syncDriveKnowledge() {
  const folderId = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_FOLDER_ID);
  if (!folderId) return "連携エラー: スクリプトプロパティに 'KNOWLEDGE_FOLDER_ID' が設定されていません。";

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sources');
    
    const existingIds = getSheetData('sources').map(row => String(row['ID']));
    let addedCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      const mimeType = file.getMimeType();
      
      if (existingIds.includes(fileId)) continue;

      let extractedText = '';
      let typeName = '';

      if (mimeType === MimeType.GOOGLE_DOCS) {
        extractedText = DocumentApp.openById(fileId).getBody().getText();
        typeName = 'Google Docs';
      } else if (mimeType === MimeType.PLAIN_TEXT || mimeType === MimeType.CSV) {
        extractedText = file.getBlob().getDataAsString();
        typeName = 'Text/CSV';
      } else if (mimeType === MimeType.PDF || mimeType === MimeType.JPEG || mimeType === MimeType.PNG) {
        try {
          const tempDoc = Drive.Files.insert({
            title: file.getName() + ' (Temp OCR)',
            mimeType: MimeType.GOOGLE_DOCS
          }, file.getBlob(), {ocr: true});
          extractedText = DocumentApp.openById(tempDoc.id).getBody().getText();
          Drive.Files.remove(tempDoc.id);
          typeName = 'PDF/Image (OCR)';
        } catch(ocrErr) {
          extractedText = '※OCR処理エラー。拡張サービス「Drive API」が有効か確認してください。詳細: ' + ocrErr.message;
          typeName = 'OCR Failed';
        }
      } else {
        continue;
      }

      if (extractedText.length > 30000) {
        extractedText = extractedText.substring(0, 30000) + '\n...(以下略)';
      }

      const row = [fileId, file.getName(), file.getUrl(), typeName, getCurrentTime(), extractedText];
      sheet.appendRow(row);
      addedCount++;
    }

    return `同期完了: 新たに ${addedCount} 件のドキュメントをAIナレッジに変換しました。`;
  } catch (e) {
    return 'Driveファイル同期中にエラーが発生しました: ' + e.message;
  }
}

// ----------------------------------------------------
// Google Workspace フル連動: カレンダー＆メール分析とChat発信
// ----------------------------------------------------
function analyzeWorkspaceAndPushChat() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  if (!apiKey) return '設定エラー: Gemini APIキーがありません。';
  
  try {
    const now = new Date();
    const past = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
    const future = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000));
    const events = CalendarApp.getDefaultCalendar().getEvents(past, future);
    
    let calendarContext = '【カレンダー予定 (過去30日〜未来7日)】\n';
    events.forEach(e => {
      calendarContext += `- ${Utilities.formatDate(e.getStartTime(), 'JST', 'MM/dd')} : ${e.getTitle()}\n`;
    });

    const threads = GmailApp.search('in:sent OR label:inbox', 0, 30);
    let mailContext = '【最近のメールやり取り要約】\n';
    threads.forEach(t => {
      const msg = t.getMessages()[0];
      mailContext += `- 件名: ${msg.getSubject()} (日付: ${Utilities.formatDate(msg.getDate(), 'JST', 'MM/dd')})\n`;
      let bodySnippet = msg.getPlainBody().substring(0, 200).replace(/\n/g, ' ');
      mailContext += `  内容: ${bodySnippet}...\n`;
    });

    const prompt = `あなたは社長/営業マンの非常に優秀な「エグゼクティブ・アシスタントAI」です。\n以下の【直近のスケジュール】と【最近のメール】を分析し、以下の3点を簡潔に（マークダウン形式で）レポートしてください。\n1. ユーザーが最近どんな業務に注力し、何を学んだか（傾向分析）\n2. 顧客（または社内）とどのような情報交換が行われているかのトレンド\n3. 今後の予定を踏まえた、明日以降の「推奨されるネクストアクション（タスク）」\n\n${calendarContext}\n\n${mailContext}`;
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    const payload = { 'contents': [{'parts': [{'text': prompt}]}] };
    const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload) };
    
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const analysisReport = json.candidates[0].content.parts[0].text;

    const webhookUrl = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_CHAT_WEBHOOK);
    if (webhookUrl) {
      const chatPayload = { 'text': '📊 *自動業務分析レポート*\n===================\n\n' + analysisReport };
      UrlFetchApp.fetch(webhookUrl, {
        'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(chatPayload), 'muteHttpExceptions': true
      });
    }

    return analysisReport;

  } catch(e) {
    return '分析中にエラーが発生しました。スクリプト実行者がGmail/Calendarの権限を許可しているか確認してください。詳細: ' + e.message;
  }
}

// ----------------------------------------------------
// AI ナレッジ検索＆新人ロープレ機能 (Gemini API 連携)
// ----------------------------------------------------

function processAIQuery(query, mode) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  if (!apiKey) {
    return 'システムエラー: Gemini APIキーが設定されていません。管理者に連絡するか、初期設定手順を確認してください。';
  }

  const products = getSheetData('products', 100); 
  const meetings = getSheetData('meetings', 30);
  const notes    = getSheetData('notes', 30);
  const sources  = getSheetData('sources', 30);

  let contextText = '【製品マスタ情報】\n';
  contextText += products.map(p => `- 製品名: ${p['製品名']} (価格: ${p['価格']}, 状態: ${p['ステータス']})\n  優位性: ${p['競合優位性']}\n  後継機種: ${p['後継機種']}, 改廃理由: ${p['改廃理由']}`).join('\n');
  
  contextText += '\n\n【最新の商談議事録】\n';
  contextText += meetings.map(m => `- [${m['登録日時']}] 顧客: ${m['顧客名']} (関連: ${m['関連基板・機種']})\n  サマリ: ${m['内容サマリ']}`).join('\n');
  
  contextText += '\n\n【社内ナレッジメモ】\n';
  contextText += notes.map(n => `- タグ[${n['タグ(Obsidianタグ等)']}] 内容: ${n['メモ内容']}`).join('\n');

  contextText += '\n\n【Driveドキュメント（仕様書・BOM・PDF等）】\n';
  contextText += sources.map(s => `- [ファイル: ${s['ファイル名']}] (${s['タイプ']})\n  内容抜粋:\n${(s['テキスト内容'] || '').substring(0, 10000)}...`).join('\n\n');

  let systemInstruction = '';
  if (mode === 'roleplay') {
    systemInstruction = `あなたは製造業企業の厳格な「顧客（購買部長や技術責任者）」です。提供された当社の【製品マスタ情報】や【最新の商談議事録】の背景情報を踏まえつつ、新人営業担当者の提案や質問に対して、あえて鋭い指摘、価格交渉、または代替品の性能に対する現実的な懸念を投げかけてください。応答は対話（ロープレ会話）形式とし、簡潔でリアリティのある口調にしてください。`;
  } else if (mode === 'clone') {
    systemInstruction = `あなたは私の「完全なデジタルクローンAI」です。提供された当社の【製品マスタ情報】、【商談議事録】、【ナレッジメモ】、【Driveドキュメント】の全てをあなた自身の過去の経験・記憶として扱い、私の思考プロセスや言葉遣い、営業のノウハウを完全に模倣して返答してください。顧客対応のドラフト作成や、戦略の壁打ち相手として最適かつ実践的なアクションを提示してください。\n\n【あなたの記憶データ】\n${contextText}`;
  } else {
    systemInstruction = `あなたは組織の知を統合した優秀な「セールス・イネーブルメントAI」です。ユーザーからの質問に対し、以下の【社内データ】のみを情報源として、最も適切で根拠に基づいた回答を作成してください。社内で解決すべき方針が見える場合はアドバイスも添えてください。ただし、提供された社内データに該当する情報が全くない場合は、推測で語らず「社内データベースには該当情報が見つかりません」と正直に回答してください。\n\n${contextText}`;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    'system_instruction': { 'parts': { 'text': systemInstruction } },
    'contents': [{ 'parts': [{ 'text': query }] }],
    'generationConfig': { 'temperature': mode === 'roleplay' ? 0.8 : 0.2 }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    let aiResponse = '';
    if (json.candidates && json.candidates.length > 0) {
      aiResponse = json.candidates[0].content.parts[0].text;
    } else {
      aiResponse = 'Gemini APIエラー: 応答の解析に失敗しました。詳細: ' + JSON.stringify(json);
    }

    appendToSheet('qalogs', [getCurrentTime(), query, aiResponse]);
    return aiResponse;
  } catch (e) {
    return '通信エラーが発生しました。ネットワーク設定またはAPI制限を確認してください: ' + e.message;
  }
}

// ============================================================
// CSV / Excel / スプレッドシート インポート機能
// ============================================================

/**
 * フロントエンドから送られた行データを一括インポートする
 * @param {string} sheetType - 'products' | 'meetings' | 'notes' | 'goals'
 * @param {string[]} headers - ヘッダー列名の配列
 * @param {string[][]} dataRows - データ行の2次元配列
 */
function importBulkData(sheetType, headers, dataRows) {
  const SHEET_CONFIG = {
    products: {
      sheetName: 'products',
      requiredCols: ['製品名'],
      buildRow: (h, r) => [
        generateId(),
        getCellValue(h, r, '製品名'),
        getCellValue(h, r, '価格'),
        getCellValue(h, r, 'ステータス') || '現行品',
        getCellValue(h, r, '改廃理由'),
        getCellValue(h, r, '後継機種'),
        getCellValue(h, r, '競合優位性')
      ]
    },
    meetings: {
      sheetName: 'meetings',
      requiredCols: ['顧客名'],
      buildRow: (h, r) => [
        generateId(),
        getCurrentTime(),
        getCellValue(h, r, '顧客名'),
        getCellValue(h, r, '関連基板・機種'),
        getCellValue(h, r, '内容サマリ'),
        getCellValue(h, r, '議事録全文')
      ]
    },
    notes: {
      sheetName: 'notes',
      requiredCols: ['メモ内容'],
      buildRow: (h, r) => [
        generateId(),
        getCurrentTime(),
        getCellValue(h, r, 'タグ(Obsidianタグ等)'),
        getCellValue(h, r, 'メモ内容')
      ]
    },
    goals: {
      sheetName: 'goals',
      requiredCols: ['目標タイトル'],
      buildRow: (h, r) => [
        generateId(),
        getCurrentTime(),
        getCellValue(h, r, '対象期間'),
        getCellValue(h, r, '目標タイトル'),
        getCellValue(h, r, '測定指標(KPI)'),
        getCellValue(h, r, '進捗(%)') || 0,
        getCellValue(h, r, '達成状況・自己評価'),
        ''
      ]
    }
  };

  const config = SHEET_CONFIG[sheetType];
  if (!config) throw new Error('不明なインポート種別: ' + sheetType);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + config.sheetName);

  // 必須列チェック
  for (const req of config.requiredCols) {
    if (!headers.includes(req)) {
      throw new Error(`必須列「${req}」が見つかりません。テンプレートを確認してください。`);
    }
  }

  let successCount = 0;
  let skipCount = 0;

  for (const row of dataRows) {
    // 空行スキップ
    if (row.every(cell => !String(cell).trim())) { skipCount++; continue; }
    try {
      const newRow = config.buildRow(headers, row);
      sheet.appendRow(newRow);
      successCount++;
    } catch(e) {
      skipCount++;
    }
  }

  return `インポート完了: ${successCount} 件を登録しました。（${skipCount} 件スキップ）`;
}

/**
 * ヘッダー配列と行配列から特定列の値を取得するヘルパー
 */
function getCellValue(headers, row, columnName) {
  const idx = headers.indexOf(columnName);
  if (idx === -1) return '';
  return String(row[idx] || '').trim();
}

/**
 * 別のスプレッドシートからデータを読み込む
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} tabName - シート名（空の場合は先頭シート）
 */
function readFromSpreadsheet(spreadsheetId, tabName) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    let sheet;
    if (tabName && tabName.trim() !== '') {
      sheet = ss.getSheetByName(tabName.trim());
      if (!sheet) throw new Error(`シート名「${tabName}」が見つかりません。スプレッドシート内のシート名を確認してください。`);
    } else {
      sheet = ss.getSheets()[0];
    }

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) {
      return { headers: [], rows: [] };
    }

    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1)
      .filter(row => row.some(cell => String(cell).trim() !== ''))
      .map(row => row.map(cell => String(cell)));

    return { headers: headers, rows: rows };
  } catch(e) {
    throw new Error('スプレッドシートへのアクセスに失敗しました: ' + e.message + ' ※スプレッドシートの共有設定または同一アカウントでのアクセスを確認してください。');
  }
}

// ============================================================
// カレンダー機能 & Google Keep 連携
// ============================================================

/**
 * カレンダー表示用のイベントデータを一括取得
 * meetings, notes, goals シートから日付情報を含めて返す
 */
function getCalendarEvents() {
  const events = [];

  // ---- meetings ----
  const meetings = getSheetData('meetings');
  meetings.forEach(m => {
    const dateStr = extractDateStr(m['登録日時'] || '');
    const calDate = m['打合せ日'] ? extractDateStr(String(m['打合せ日'])) : dateStr;
    if (!calDate) return;
    events.push({
      id:      m['ID'] || generateId(),
      type:    'meeting',
      date:    calDate,
      time:    m['開始時刻'] || '',
      endtime: m['終了時刻'] || '',
      client:  m['顧客名'] || '',
      title:   m['件名'] || '',
      related: m['関連基板・機種'] || '',
      summary: m['内容サマリ'] || '',
      fulltext:m['議事録全文'] || ''
    });
  });

  // ---- notes ----
  const notes = getSheetData('notes');
  notes.forEach(n => {
    const dateStr = extractDateStr(n['登録日時'] || '');
    const calDate = n['対象日'] ? extractDateStr(String(n['対象日'])) : dateStr;
    if (!calDate) return;
    events.push({
      id:      n['ID'] || generateId(),
      type:    'note',
      date:    calDate,
      time:    n['時刻'] || '',
      endtime: '',
      tags:    n['タグ(Obsidianタグ等)'] || '',
      content: n['メモ内容'] || ''
    });
  });

  // ---- goals ----
  const goals = getSheetData('goals');
  goals.forEach(g => {
    const dateStr = extractDateStr(g['登録日時'] || '');
    if (!dateStr) return;
    events.push({
      id:      g['ID'] || generateId(),
      type:    'goal',
      date:    dateStr,
      time:    '',
      title:   g['目標タイトル'] || '',
      progress:g['進捗(%)'] || 0
    });
  });

  return events;
}

/**
 * 日時文字列から YYYY-MM-DD を抽出するユーティリティ
 */
function extractDateStr(str) {
  if (!str) return '';
  const s = String(str);
  // YYYY/MM/DD または YYYY-MM-DD
  const m1 = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m1) return `${m1[1]}-${m1[2].padStart(2,'0')}-${m1[3].padStart(2,'0')}`;
  // Date オブジェクト系
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  } catch(e) {}
  return '';
}

/**
 * 打合せ日付・時刻付きで商談議事録を保存
 */
function saveMeetingWithDate(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('meetings');

  // meetingsシートに打合せ日・開始/終了時刻列を動的追加（なければ）
  ensureColumns('meetings', ['打合せ日', '件名', '開始時刻', '終了時刻']);

  if (data.id) {
    // 既存レコードの更新
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        const colMap = {};
        headers.forEach((h, ci) => colMap[h] = ci + 1);
        if (colMap['顧客名'])        sheet.getRange(i+1, colMap['顧客名']).setValue(data.client);
        if (colMap['件名'])          sheet.getRange(i+1, colMap['件名']).setValue(data.title || '');
        if (colMap['打合せ日'])      sheet.getRange(i+1, colMap['打合せ日']).setValue(data.date);
        if (colMap['開始時刻'])      sheet.getRange(i+1, colMap['開始時刻']).setValue(data.time || '');
        if (colMap['終了時刻'])      sheet.getRange(i+1, colMap['終了時刻']).setValue(data.endtime || '');
        if (colMap['関連基板・機種'])sheet.getRange(i+1, colMap['関連基板・機種']).setValue(data.related || '');
        if (colMap['内容サマリ'])    sheet.getRange(i+1, colMap['内容サマリ']).setValue(data.summary || '');
        if (colMap['議事録全文'])    sheet.getRange(i+1, colMap['議事録全文']).setValue(data.fulltext || '');
        return '議事録を更新しました。';
      }
    }
  }

  // 新規追加
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = {};
  headers.forEach((h, ci) => colMap[h] = ci);

  // 基本7列 + 拡張列の順で行データを構築
  const baseRow = [
    generateId(), getCurrentTime(),
    data.client || '',
    data.related || '',
    data.summary || '',
    data.fulltext || ''
  ];

  // 拡張列は別途 setValue で設定
  sheet.appendRow(baseRow);
  const newRow = sheet.getLastRow();
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  hdr.forEach((h, ci) => {
    if (h === '打合せ日')      sheet.getRange(newRow, ci+1).setValue(data.date || '');
    if (h === '件名')          sheet.getRange(newRow, ci+1).setValue(data.title || '');
    if (h === '開始時刻')      sheet.getRange(newRow, ci+1).setValue(data.time || '');
    if (h === '終了時刻')      sheet.getRange(newRow, ci+1).setValue(data.endtime || '');
  });

  return '打合せ・議事録を登録しました。';
}

/**
 * 対象日・時刻付きでメモを保存
 */
function saveNoteWithDate(data) {
  ensureColumns('notes', ['対象日', '時刻']);

  if (data.id) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('notes');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        const colMap = {};
        headers.forEach((h, ci) => colMap[h] = ci + 1);
        if (colMap['タグ(Obsidianタグ等)']) sheet.getRange(i+1, colMap['タグ(Obsidianタグ等)']).setValue(data.tags || '');
        if (colMap['メモ内容'])            sheet.getRange(i+1, colMap['メモ内容']).setValue(data.content || '');
        if (colMap['対象日'])              sheet.getRange(i+1, colMap['対象日']).setValue(data.date || '');
        if (colMap['時刻'])                sheet.getRange(i+1, colMap['時刻']).setValue(data.time || '');
        return 'メモを更新しました。';
      }
    }
  }

  appendToSheet('notes', [generateId(), getCurrentTime(), data.tags || '', data.content || '']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('notes');
  const newRow = sheet.getLastRow();
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  hdr.forEach((h, ci) => {
    if (h === '対象日') sheet.getRange(newRow, ci+1).setValue(data.date || '');
    if (h === '時刻')   sheet.getRange(newRow, ci+1).setValue(data.time || '');
  });

  return 'メモを登録しました。';
}

/**
 * Google Keep にメモを作成する
 * ※ Keep API は現在一般公開されていないため、
 *    Googleが提供する KeepService (Apps Script 拡張サービス) を使用。
 *    利用不可環境では notesシートへの保存にフォールバック。
 */
function saveToKeep(data) {
  // ---- Keep API 試行 ----
  try {
    if (typeof Keep !== 'undefined') {
      const note = Keep.newNote();
      note.title = data.title || '';
      const body = Keep.newTextContent();
      body.text = data.body || '';
      note.body = body;

      // 色設定
      const colorMap = {
        'DEFAULT': Keep.ColorValue.DEFAULT,
        'YELLOW':  Keep.ColorValue.YELLOW,
        'GREEN':   Keep.ColorValue.GREEN,
        'BLUE':    Keep.ColorValue.BLUE,
        'RED':     Keep.ColorValue.RED,
        'PURPLE':  Keep.ColorValue.PURPLE
      };
      note.color = colorMap[data.color] || Keep.ColorValue.DEFAULT;

      Keep.Notes.create(note);

      // スプレッドシートにも記録
      _saveKeepRecord(data, true);
      return 'Google Keep にメモを保存しました。';
    }
  } catch(e) {
    // Keep API が利用できない場合はフォールバック
  }

  // ---- フォールバック: スプレッドシートに保存 ----
  _saveKeepRecord(data, false);
  return 'ナレッジメモ（Keep同等）として登録しました。※Google Keep APIをApps Scriptで有効化すると直接Keepに保存されます。';
}

function _saveKeepRecord(data, sentToKeep) {
  ensureColumns('notes', ['対象日', '時刻', 'Keepラベル', 'Keepカラー', 'Keep送信済み']);
  appendToSheet('notes', [
    generateId(), getCurrentTime(),
    '#keep ' + (data.labels || ''),
    (data.title ? '【' + data.title + '】\n' : '') + (data.body || '')
  ]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('notes');
  const newRow = sheet.getLastRow();
  const hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  hdr.forEach((h, ci) => {
    if (h === '対象日')      sheet.getRange(newRow, ci+1).setValue(data.date || '');
    if (h === 'Keepラベル')  sheet.getRange(newRow, ci+1).setValue(data.labels || '');
    if (h === 'Keepカラー')  sheet.getRange(newRow, ci+1).setValue(data.color || 'DEFAULT');
    if (h === 'Keep送信済み')sheet.getRange(newRow, ci+1).setValue(sentToKeep ? '済' : '未');
  });
}

/**
 * Google Keep からメモ一覧を取得
 */
function getKeepNotes() {
  const results = [];

  // ---- Keep API 試行 ----
  try {
    if (typeof Keep !== 'undefined') {
      const notes = Keep.Notes.list({ pageSize: 50 });
      (notes.notes || []).forEach(n => {
        results.push({
          id:     n.name || generateId(),
          title:  n.title || '',
          body:   n.body && n.body.text ? n.body.text.text : '',
          color:  n.color || 'DEFAULT',
          labels: (n.labels || []).map(l => l.name || '')
        });
      });
      return results;
    }
  } catch(e) {}

  // ---- フォールバック: notesシートから#keepタグのものを返す ----
  const notes = getSheetData('notes');
  notes.filter(n => String(n['タグ(Obsidianタグ等)']).includes('#keep') || String(n['タグ(Obsidianタグ等)']).includes('keep'))
    .reverse().slice(0, 30)
    .forEach(n => {
      results.push({
        id:    String(n['ID']),
        title: (String(n['メモ内容'] || '')).startsWith('【') ? String(n['メモ内容']).match(/【(.+?)】/)?.[1] || '' : '',
        body:  String(n['メモ内容'] || ''),
        color: n['Keepカラー'] || 'DEFAULT',
        labels: String(n['Keepラベル'] || '').split(',').map(s=>s.trim()).filter(Boolean)
      });
    });
  return results;
}

/**
 * シートに指定列が存在しない場合は末尾に追加する
 */
function ensureColumns(sheetName, columnNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  columnNames.forEach(col => {
    if (!headers.includes(col)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    }
  });
}
