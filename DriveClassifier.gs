// ==========================================
// DriveClassifier.gs  (完全版)
// 画像フォルダ構成①②両対応
// ==========================================

const CLASSIFIER_RULES = {
  pcb_bom:      { patterns: ['BOM', '部品表', '構成表', 'parts_list', 'bill_of_material', 'partslist', '部品リスト', '実装表'],    label: '部品表(BOM)',    icon: '📦', pcbType: 'bom'    },
  pcb_board:    { patterns: ['基板仕様', '基板設計', '回路設計', 'schematic', 'PCB_spec', '基板Rev', '基板改版', '基板構成'],       label: '基板仕様書',     icon: '🔧', pcbType: 'board'  },
  pcb_model:    { patterns: ['機種仕様', '製品仕様', '機種概要', 'model_spec', '機番', '型式説明'],                                 label: '機種仕様書',     icon: '🏭', pcbType: 'model'  },
  pcb_revision: { patterns: ['改版履歴', '変更履歴', 'revision_history', '改訂履歴', 'ECN', '設計変更'],                           label: '改版履歴',       icon: '📝', pcbType: 'revision'},
  model_spec:   { patterns: ['製品仕様', '機種仕様', 'spec', '仕様書'],                       label: '機種製品仕様書', icon: '📋' },
  board_design: { patterns: ['基板設計', '回路設計', '基板仕様', 'schematic', 'PCB'],          label: '基板設計仕様書', icon: '🔧' },
  bom:          { patterns: ['BOM', '構成表', 'bill_of_material', '部品表', 'parts_list'],     label: '構成表(BOM)',   icon: '📦' },
  quote:        { patterns: ['見積', 'quote', 'QUOTE', '見積書', 'estimate'],                  label: '見積書',        icon: '💴' },
  drawing:      { patterns: ['図面', 'drawing', 'DWG', 'CAD', '外形図'],                      label: '図面',          icon: '📐' },
  manual:       { patterns: ['マニュアル', 'manual', '操作説明', '取扱説明'],                  label: 'マニュアル',    icon: '📖' },
  test:         { patterns: ['試験', 'テスト', 'test', '検査', '評価'],                        label: '試験・検査書',  icon: '🧪' },
};

// ==========================================
// メイン: フォルダタイプ自動判定＆スキャン
// ==========================================

/**
 * @param {string} rootFolderId - ルートフォルダID
 * @param {number} folderType   - 1: 見積書→会社→月  / 2: 会社→見積書→月
 */
function scanAndClassifyDriveFolder(rootFolderId, folderType) {
  if (!rootFolderId) return { status: 'error', message: 'フォルダIDが未入力です。' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureClassifierSheets(ss);

  try {
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const results    = { classified: [], unclassified: [], companies: [] };

    if (String(folderType) === '2') {
      // ===== 構成② 会社フォルダ が直下にある =====
      // ルート直下: [本社], [○○支社], ...
      scanType2(ss, rootFolder, results);
    } else {
      // ===== 構成① 見積書(ルート) → 会社 → 月 =====
      scanType1(ss, rootFolder, results);
    }

    // PCBナレッジ自動インポート試行
    let pcbImported = 0;
    results.classified.forEach(doc => {
      try {
        const rule = Object.values(CLASSIFIER_RULES).find(r => r.label === doc.type);
        if (rule && rule.pcbType === 'bom' && doc.fileId) {
          const ctx = { modelId: '', boardId: '' };
          const res = importFromDriveFile(doc.fileId, null, 'bom', ctx);
          if (res && res.imported) pcbImported += res.imported;
        }
      } catch(e) {}
    });

    return {
      status: 'success',
      classified:       results.classified.length,
      unclassified:     results.unclassified.length,
      pcbImported:      pcbImported,
      details:          results.classified,
      unclassifiedFiles: results.unclassified,
      companies:        results.companies,
      message: `✅ 分類完了: ${results.classified.length}件 自動分類 / ${results.unclassified.length}件 未分類` + (pcbImported > 0 ? ` / 基板ナレッジ ${pcbImported}件インポート` : '')
    };
  } catch(e) {
    return { status: 'error', message: 'スキャンエラー: ' + e.message };
  }
}

// --------------------------------------------------
// 構成①: [見積書ルート] → [会社] → [年月] → ファイル
// --------------------------------------------------
function scanType1(ss, rootFolder, results) {
  const companyFolders = rootFolder.getFolders();

  while (companyFolders.hasNext()) {
    const companyFolder = companyFolders.next();
    const companyName   = companyFolder.getName();

    // 会社レコード作成
    const companyId = getOrCreateCompany(ss, companyName, companyFolder.getId(), companyFolder.getUrl(), 'type1');
    results.companies.push(companyName);

    // 月別フォルダを処理
    const monthFolders = companyFolder.getFolders();
    while (monthFolders.hasNext()) {
      const monthFolder = monthFolders.next();
      const yearMonth   = extractYearMonth(monthFolder.getName());
      processMonthFolder(ss, monthFolder, companyId, companyName, yearMonth, results);
    }

    // 会社フォルダ直下のファイルも処理
    processFilesInFolder(ss, companyFolder, companyId, companyName, '', results);
  }
}

// --------------------------------------------------
// 構成②: [会社] → [見積書] → [年月] → ファイル
// --------------------------------------------------
function scanType2(ss, rootFolder, results) {
  const companyFolders = rootFolder.getFolders();

  while (companyFolders.hasNext()) {
    const companyFolder = companyFolders.next();
    const companyName   = companyFolder.getName();

    const companyId = getOrCreateCompany(ss, companyName, companyFolder.getId(), companyFolder.getUrl(), 'type2');
    results.companies.push(companyName);

    // 会社フォルダ直下のサブフォルダを走査（見積書、その他）
    const subFolders = companyFolder.getFolders();
    while (subFolders.hasNext()) {
      const sub     = subFolders.next();
      const subName = sub.getName();

      if (isQuoteFolder(subName)) {
        // 見積書フォルダ → 月別を処理
        const monthFolders = sub.getFolders();
        while (monthFolders.hasNext()) {
          const monthFolder = monthFolders.next();
          const yearMonth   = extractYearMonth(monthFolder.getName());
          processMonthFolder(ss, monthFolder, companyId, companyName, yearMonth, results);
        }
        // 見積書フォルダ直下ファイル
        processFilesInFolder(ss, sub, companyId, companyName, '', results);
      } else {
        // 見積書以外のサブフォルダ（機種フォルダ等）
        processFilesInFolder(ss, sub, companyId, companyName, subName, results);
      }
    }

    // 会社直下ファイル
    processFilesInFolder(ss, companyFolder, companyId, companyName, '', results);
  }
}

// --------------------------------------------------
// 月別フォルダ内のファイルを処理
// --------------------------------------------------
function processMonthFolder(ss, monthFolder, companyId, companyName, yearMonth, results) {
  const files = monthFolder.getFiles();
  while (files.hasNext()) {
    const file           = files.next();
    const classification = classifyFile(file.getName(), file.getMimeType());

    if (classification) {
      saveClassifiedDoc(ss, file, classification, companyId, companyName, yearMonth);
      results.classified.push({
        name:    file.getName(),
        type:    classification.label,
        icon:    classification.icon,
        company: companyName,
        month:   yearMonth
      });
    } else {
      saveUnclassifiedDoc(ss, file, companyId);
      results.unclassified.push({ name: file.getName(), url: file.getUrl() });
    }
  }
}

// --------------------------------------------------
// フォルダ直下ファイルを処理（月別以外）
// --------------------------------------------------
function processFilesInFolder(ss, folder, companyId, companyName, context, results) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file           = files.next();
    const classification = classifyFile(file.getName(), file.getMimeType());

    if (classification) {
      saveClassifiedDoc(ss, file, classification, companyId, companyName, context);
      results.classified.push({
        name:    file.getName(),
        type:    classification.label,
        icon:    classification.icon,
        company: companyName,
        month:   context
      });
    } else {
      saveUnclassifiedDoc(ss, file, companyId);
      results.unclassified.push({ name: file.getName(), url: file.getUrl() });
    }
  }
}

// ==========================================
// 自動分類ロジック
// ==========================================
function classifyFile(fileName, mimeType) {
  const lowerName = fileName.toLowerCase().replace(/[-_\s]/g, '');
  for (const [key, rule] of Object.entries(CLASSIFIER_RULES)) {
    for (const pattern of rule.patterns) {
      if (lowerName.includes(pattern.toLowerCase().replace(/[-_\s]/g, ''))) {
        return { key, label: rule.label, icon: rule.icon };
      }
    }
  }
  // スプレッドシート/CSVはBOMの可能性が高い
  if (mimeType === MimeType.GOOGLE_SHEETS || /\.xlsx?$/i.test(fileName) || mimeType === 'text/csv') {
    return CLASSIFIER_RULES.pcb_bom || CLASSIFIER_RULES.bom;
  }
  return null;
}

function isQuoteFolder(name) {
  return ['見積', 'quote', '見積書'].some(k => name.toLowerCase().includes(k.toLowerCase()));
}

// 「2025年01月」「2025-01」「202501」などから "2025年01月" 形式に正規化
function extractYearMonth(folderName) {
  const m =
    folderName.match(/(\d{4})[年\-\/](\d{1,2})[月]?/) ||
    folderName.match(/(\d{4})(\d{2})/);
  if (m) return `${m[1]}年${String(m[2]).padStart(2,'0')}月`;
  return folderName; // 判定できない場合そのまま
}

// ==========================================
// スプレッドシートへの書き込み
// ==========================================
function getOrCreateCompany(ss, name, folderId, folderUrl, type) {
  const sheet = ss.getSheetByName('companies');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === folderId) return data[i][0];
  }
  const id = generateId();
  sheet.appendRow([id, name, folderId, folderUrl, type, getCurrentTime()]);
  return id;
}

function saveClassifiedDoc(ss, file, classification, companyId, companyName, yearMonth) {
  const sheet = ss.getSheetByName('doc_hierarchy');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === file.getId()) return; // 重複スキップ
  }
  // 見積書は quotesシートにも転記
  if (classification.key === 'quote') {
    saveToQuoteSheet(ss, file, companyId, companyName, yearMonth);
  }
  sheet.appendRow([
    generateId(), companyId, companyName, yearMonth,
    classification.label, file.getName(),
    file.getId(), file.getUrl(), getCurrentTime()
  ]);
}

function saveUnclassifiedDoc(ss, file, companyId) {
  const sheet = ss.getSheetByName('doc_hierarchy');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === file.getId()) return;
  }
  sheet.appendRow([
    generateId(), companyId, '', '', '【未分類】',
    file.getName(), file.getId(), file.getUrl(), getCurrentTime()
  ]);
}

function saveToQuoteSheet(ss, file, companyId, companyName, yearMonth) {
  const sheet = ss.getSheetByName('quotes');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === file.getId()) return;
  }
  const m        = file.getName().match(/([A-Z]+-?\d+)/);
  const quoteNum = m ? m[1] : file.getName().replace(/\.\w+$/, '');
  sheet.appendRow([
    generateId(), companyId, companyName, quoteNum,
    yearMonth, file.getId(), file.getUrl(), '未確認', getCurrentTime()
  ]);
}

// ==========================================
// シート自動生成
// ==========================================
function ensureClassifierSheets(ss) {
  const required = {
    'companies':     ['会社ID', '会社名', 'DriveフォルダID', 'Drive URL', '構成タイプ', '登録日時'],
    'quotes':        ['見積ID', '親会社ID', '会社名', '見積番号', '年月', 'DriveファイルID', 'Drive URL', 'ステータス', '登録日時'],
    'doc_hierarchy': ['ドキュメントID', '親会社ID', '会社名', '年月', 'ファイル種別', 'ファイル名', 'DriveファイルID', 'Drive URL', '登録日時'],
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
 * 会社別・月別ツリーデータを返す
 */
function getModelTree() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureClassifierSheets(ss);

  const companies = getSheetData('companies');
  const docs      = getSheetData('doc_hierarchy');
  const quotes    = getSheetData('quotes');

  return companies.map(company => {
    const cId       = company['会社ID'];
    const compDocs  = docs.filter(d => d['親会社ID'] === cId);
    const compQuotes = quotes.filter(q => q['親会社ID'] === cId);

    // 年月ごとにグループ化
    const monthMap = {};
    compDocs.forEach(d => {
      const ym = d['年月'] || '（月不明）';
      if (!monthMap[ym]) monthMap[ym] = [];
      monthMap[ym].push(d);
    });

    return {
      id:      cId,
      name:    company['会社名'],
      url:     company['Drive URL'],
      type:    company['構成タイプ'],
      months:  Object.entries(monthMap)
                 .sort((a, b) => b[0].localeCompare(a[0])) // 新しい月順
                 .map(([ym, files]) => ({ yearMonth: ym, files })),
      quotes:  compQuotes,
      totalDocs: compDocs.length
    };
  });
}

function reclassifyDocument(docId, newType) {
  return updateRowData('doc_hierarchy', docId, 'ファイル種別', newType);
}

function validateDriveFolderId(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    return { valid: true, name: folder.getName(), url: folder.getUrl() };
  } catch(e) {
    return { valid: false, message: 'フォルダが見つかりません。IDを確認してください。' };
  }
}
