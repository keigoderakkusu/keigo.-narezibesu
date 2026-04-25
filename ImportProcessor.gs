// ==========================================
// ImportProcessor.gs
// CSV / PDF / スプレッドシート インポート処理
// ==========================================

// --------------------------------------------------
// メインエントリ: CSVテキスト → PCB階層へ登録
// --------------------------------------------------
function importCSV(csvText, importType, context) {
  try {
    const rows = _parseCSV(csvText);
    if (rows.length < 2) return { status: 'error', message: 'CSVにデータ行がありません' };
    const headers  = rows[0].map(h => h.trim());
    const dataRows = rows.slice(1).filter(r => r.some(c => String(c).trim() !== ''));
    const type     = importType === 'auto' ? _autoDetectType(headers) : importType;

    switch (type) {
      case 'bom':    return _importBOM(headers, dataRows, context || {});
      case 'models': return _importModels(headers, dataRows);
      case 'boards': return _importBoards(headers, dataRows, context || {});
      default: return { status: 'error', message: '不明なタイプ: ' + type };
    }
  } catch(e) {
    return { status: 'error', message: 'CSVエラー: ' + e.message };
  }
}

// --------------------------------------------------
// Driveファイル(スプレッドシート / CSV)からインポート
// --------------------------------------------------
function importFromDriveFile(fileId, sheetName, importType, context) {
  try {
    const file     = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    let csvText;

    if (mimeType === MimeType.GOOGLE_SHEETS) {
      const ss    = SpreadsheetApp.openById(fileId);
      const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
      if (!sheet) return { status: 'error', message: `シート "${sheetName}" が見つかりません` };
      const data = sheet.getDataRange().getValues();
      csvText = data.map(row => row.map(c => {
        const s = String(c);
        return s.includes(',') ? `"${s.replace(/"/g, '""')}"` : s;
      }).join(',')).join('\n');
    } else if (mimeType === 'text/csv' || /\.csv$/i.test(file.getName())) {
      csvText = file.getBlob().getDataAsString('UTF-8');
      // BOM除去
      if (csvText.charCodeAt(0) === 0xFEFF) csvText = csvText.slice(1);
    } else if (mimeType === MimeType.MICROSOFT_EXCEL || /\.xlsx?$/i.test(file.getName())) {
      return { status: 'error', message: 'ExcelファイルはGoogleスプレッドシートに変換してからインポートしてください。\nDriveでファイルを右クリック→「アプリで開く」→「Googleスプレッドシート」で変換できます。' };
    } else {
      return { status: 'error', message: '対応形式: GoogleスプレッドシートまたはCSVファイル' };
    }

    return importCSV(csvText, importType, context);
  } catch(e) {
    return { status: 'error', message: 'Driveアクセスエラー: ' + e.message };
  }
}

// --------------------------------------------------
// PDFファイルからテキスト抽出 → Geminiで部品情報解析
// --------------------------------------------------
function importPDFtoBOM(fileId, context) {
  try {
    const file = DriveApp.getFileById(fileId);
    if (file.getMimeType() !== MimeType.PDF) {
      return { status: 'error', message: 'PDFファイルを指定してください' };
    }

    // PDFをGoogleドキュメントに一時変換してテキスト取得
    let text = '';
    let tmpDocId = null;
    try {
      const resource = { name: '__tmp_pcb_pdf_' + fileId, mimeType: MimeType.GOOGLE_DOCS };
      const copied   = Drive.Files.copy(resource, fileId);
      tmpDocId = copied.id;
      const doc = DocumentApp.openById(tmpDocId);
      text = doc.getBody().getText();
    } finally {
      if (tmpDocId) {
        try { DriveApp.getFileById(tmpDocId).setTrashed(true); } catch(e) {}
      }
    }

    if (!text.trim()) return { status: 'error', message: 'PDFからテキストを抽出できませんでした（画像PDFの場合はCSVでインポートしてください）' };

    // Geminiで部品表解析
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { status: 'error', message: 'Gemini API Keyが未設定です' };

    const prompt = `以下のPDF文書から部品表(BOM)の情報を抽出し、JSON配列で返してください。
各要素のフォーマット: {"refNo":"参照番号","partNumber":"部品番号","partName":"部品名","manufacturer":"メーカー","spec":"仕様・規格","quantity":"数量","unit":"単位"}
部品表が見つからない場合は空配列 [] のみ返してください。JSON以外のテキストは不要です。

文書テキスト:
${text.substring(0, 8000)}`;

    const resp = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
      {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1 }
        })
      }
    );
    const result  = JSON.parse(resp.getContentText());
    const raw     = result?.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
    const match   = raw.match(/\[[\s\S]*\]/);
    const partsArr = match ? JSON.parse(match[0]) : [];

    let imported = 0;
    partsArr.forEach(p => {
      if (p.partName || p.partNumber) {
        savePart({ ...p, boardId: (context || {}).boardId || '', configId: (context || {}).configId || '' });
        imported++;
      }
    });

    // sourcesシートに登録
    appendToSheet('sources', [generateId(), file.getName(), file.getUrl(), 'PDF(PCB)', getCurrentTime(), text.substring(0, 300)]);

    return { status: 'ok', imported, message: `PDF解析完了: ${imported}件の部品を登録しました (${file.getName()})` };
  } catch(e) {
    return { status: 'error', message: 'PDF処理エラー: ' + e.message };
  }
}

// --------------------------------------------------
// Drive上のフォルダを自動スキャンしてPCB階層に分類
// --------------------------------------------------
function scanFolderAndClassifyPCB(folderId) {
  setupPCBSheets();
  const folder = DriveApp.getFolderById(folderId);
  const results = { imported: [], errors: [] };

  _walkFolderForPCB(folder, results, { modelHint: '', boardHint: '' });

  return {
    status: 'ok',
    imported: results.imported.length,
    errors: results.errors.length,
    details: results.imported,
    message: `スキャン完了: ${results.imported.length}件処理 / ${results.errors.length}件エラー`
  };
}

function _walkFolderForPCB(folder, results, hints) {
  const folderName = folder.getName();
  const newHints   = _extractPCBHints(folderName, hints);

  // ファイル処理
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      const classified = _classifyAndImportPCBFile(file, newHints);
      if (classified) results.imported.push({ name: file.getName(), type: classified });
    } catch(e) {
      results.errors.push({ name: file.getName(), error: e.message });
    }
  }

  // サブフォルダ再帰
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    _walkFolderForPCB(subFolders.next(), results, newHints);
  }
}

function _extractPCBHints(folderName, parentHints) {
  const hints = { ...parentHints };
  // 機種名ヒント
  if (/機種|model|CR[A-Z]|PA[A-Z]/i.test(folderName)) hints.modelHint = folderName;
  // 基板名ヒント
  if (/基板|board|PCB|制御|主基/i.test(folderName)) hints.boardHint = folderName;
  return hints;
}

function _classifyAndImportPCBFile(file, hints) {
  const name     = file.getName().toLowerCase();
  const mimeType = file.getMimeType();

  // BOM/部品表ファイル
  if (/bom|部品表|parts?.?list|構成表/.test(name)) {
    if (mimeType === MimeType.GOOGLE_SHEETS || /\.xlsx?$/i.test(name) || /csv/i.test(mimeType)) {
      const context = _resolveOrCreateContext(hints);
      importFromDriveFile(file.getId(), null, 'bom', context);
      return 'BOM';
    }
    if (mimeType === MimeType.PDF) {
      const context = _resolveOrCreateContext(hints);
      importPDFtoBOM(file.getId(), context);
      return 'BOM(PDF)';
    }
  }

  // 機種仕様書
  if (/機種仕様|model.?spec|製品仕様/.test(name)) {
    if (hints.modelHint) return null; // すでに機種フォルダ内
    saveModel({ name: _cleanFileName(file.getName()), notes: file.getUrl() });
    return '機種';
  }

  // 基板仕様書
  if (/基板仕様|board.?spec|回路設計|schematic/.test(name)) {
    const context = _resolveOrCreateContext(hints);
    saveBoard({ modelId: context.modelId || '', name: _cleanFileName(file.getName()), notes: file.getUrl() });
    return '基板';
  }

  return null;
}

function _resolveOrCreateContext(hints) {
  const context = {};

  if (hints.modelHint) {
    const models = getSheetData('pcb_models');
    const found  = models.find(m => m['機種名'].includes(hints.modelHint) || hints.modelHint.includes(m['機種名']));
    if (found) {
      context.modelId = found['機種ID'];
    } else {
      const res = saveModel({ name: hints.modelHint });
      context.modelId = res.id;
    }
  }

  if (hints.boardHint && context.modelId) {
    const boards = getSheetData('pcb_boards');
    const found  = boards.find(b => b['基板名'].includes(hints.boardHint) || hints.boardHint.includes(b['基板名']));
    if (found) {
      context.boardId = found['基板ID'];
    } else {
      const res = saveBoard({ modelId: context.modelId, name: hints.boardHint });
      context.boardId = res.id;
    }
  }

  return context;
}

function _cleanFileName(name) {
  return name.replace(/\.[^.]+$/, '').replace(/[-_]/g, ' ').trim();
}

// --------------------------------------------------
// BOM インポート内部処理
// --------------------------------------------------
function _importBOM(headers, rows, context) {
  setupPCBSheets();
  const COL_MAP = {
    refNo:        ['参照番号','RefNo','Ref','Reference','レファレンス','参照','符号'],
    partNumber:   ['部品番号','PartNumber','Part No','品番','P/N','PN','型番'],
    partName:     ['部品名','PartName','Name','品名','部品名称','名称'],
    manufacturer: ['メーカー','Maker','Manufacturer','Mfr','メーカ','製造元','製造者'],
    spec:         ['仕様','Spec','Specification','Value','規格','値','特性'],
    quantity:     ['数量','Qty','Quantity','個数','Q\'ty','個'],
    unit:         ['単位','Unit'],
    notes:        ['備考','Notes','Remark','メモ','コメント'],
  };

  let imported = 0;
  rows.forEach(row => {
    const data = { boardId: context.boardId || '', configId: context.configId || '' };
    for (const [field, candidates] of Object.entries(COL_MAP)) {
      for (const c of candidates) {
        const idx = headers.findIndex(h =>
          h.replace(/[\s\-\.\(\)]/g, '').toLowerCase() === c.replace(/[\s\-\.\(\)]/g, '').toLowerCase()
        );
        if (idx !== -1 && row[idx] !== undefined && String(row[idx]).trim()) {
          data[field] = String(row[idx]).trim();
          break;
        }
      }
    }
    if (data.partName || data.partNumber) {
      savePart(data);
      imported++;
    }
  });
  return { status: 'ok', imported, message: `${imported}件の部品をインポートしました` };
}

// --------------------------------------------------
// 機種マスタ インポート
// --------------------------------------------------
function _importModels(headers, rows) {
  setupPCBSheets();
  const COL_MAP = {
    code:        ['機種コード','ModelCode','Code','コード','型式'],
    name:        ['機種名','ModelName','Name','機種','名称'],
    series:      ['シリーズ','Series'],
    status:      ['ステータス','Status','状態'],
    description: ['説明','Description','概要'],
  };
  let imported = 0;
  rows.forEach(row => {
    const data = {};
    for (const [field, candidates] of Object.entries(COL_MAP)) {
      for (const c of candidates) {
        const idx = headers.findIndex(h =>
          h.replace(/[\s\-\.]/g, '').toLowerCase() === c.replace(/[\s\-\.]/g, '').toLowerCase()
        );
        if (idx !== -1 && row[idx] !== undefined && String(row[idx]).trim()) {
          data[field] = String(row[idx]).trim();
          break;
        }
      }
    }
    if (data.name) { saveModel(data); imported++; }
  });
  return { status: 'ok', imported, message: `${imported}件の機種をインポートしました` };
}

// --------------------------------------------------
// 基板マスタ インポート
// --------------------------------------------------
function _importBoards(headers, rows, context) {
  setupPCBSheets();
  const COL_MAP = {
    code:           ['基板コード','BoardCode','Code','コード','型式'],
    name:           ['基板名','BoardName','Name','基板','名称'],
    revision:       ['改版番号','Rev','Revision','改版','リビジョン','版数'],
    revisionReason: ['改版理由','RevReason','変更理由','理由','変更内容'],
    description:    ['説明','Description','概要'],
  };
  let imported = 0;
  rows.forEach(row => {
    const data = { modelId: context.modelId || '' };
    for (const [field, candidates] of Object.entries(COL_MAP)) {
      for (const c of candidates) {
        const idx = headers.findIndex(h =>
          h.replace(/[\s\-\.]/g, '').toLowerCase() === c.replace(/[\s\-\.]/g, '').toLowerCase()
        );
        if (idx !== -1 && row[idx] !== undefined && String(row[idx]).trim()) {
          data[field] = String(row[idx]).trim();
          break;
        }
      }
    }
    if (data.name) { saveBoard(data); imported++; }
  });
  return { status: 'ok', imported, message: `${imported}件の基板をインポートしました` };
}

// --------------------------------------------------
// CSVパーサー (RFC 4180 準拠)
// --------------------------------------------------
function _parseCSV(text) {
  const rows  = [];
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  for (const line of lines) {
    if (!line.trim()) continue;
    const cells = [];
    let inQ = false, cell = '';
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQ && line[i+1] === '"') { cell += '"'; i++; }
        else { inQ = !inQ; }
      } else if (ch === ',' && !inQ) {
        cells.push(cell); cell = '';
      } else {
        cell += ch;
      }
    }
    cells.push(cell);
    rows.push(cells);
  }
  return rows;
}

// --------------------------------------------------
// CSVタイプ自動判定
// --------------------------------------------------
function _autoDetectType(headers) {
  const h = headers.join(' ').toLowerCase();
  if (/部品番号|part.?num|品番|bom|ref(erence)?\.?no?\.?|部品名/.test(h)) return 'bom';
  if (/機種コード|machine|機種名|model.?name/.test(h))                       return 'models';
  if (/基板コード|board.?code|改版|revision/.test(h))                        return 'boards';
  return 'bom';
}
