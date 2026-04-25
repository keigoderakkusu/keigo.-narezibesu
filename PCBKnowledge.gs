// ==========================================
// PCBKnowledge.gs
// 基板ナレッジ階層管理
// 機種 → 基板 → 基板構成 → 部品(BOM)
// ==========================================

const PCB_SHEETS = {
  models:  { name: 'pcb_models',  headers: ['機種ID','機種コード','機種名','シリーズ','ステータス','説明','メモ','登録日時'] },
  boards:  { name: 'pcb_boards',  headers: ['基板ID','機種ID','基板コード','基板名','改版番号','改版理由','説明','メモ','登録日時'] },
  configs: { name: 'pcb_configs', headers: ['構成ID','基板ID','構成名','構成バージョン','層数','材質','メモ','登録日時'] },
  bom:     { name: 'pcb_bom',     headers: ['部品ID','構成ID','基板ID','参照番号','部品番号','部品名','メーカー','仕様','数量','単位','リードタイム','メモ'] },
};

// --------------------------------------------------
// シート初期化
// --------------------------------------------------
function setupPCBSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const [, cfg] of Object.entries(PCB_SHEETS)) {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) sheet = ss.insertSheet(cfg.name);
    if (sheet.getLastRow() === 0) {
      const range = sheet.getRange(1, 1, 1, cfg.headers.length);
      range.setValues([cfg.headers]);
      range.setFontWeight('bold');
      range.setBackground('#0f172a');
      range.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
}

// --------------------------------------------------
// CRUD: 機種
// --------------------------------------------------
function saveModel(data) {
  setupPCBSheets();
  const id = data.id || generateId();
  if (data.id) {
    _updatePCBRow('pcb_models', id, {
      '機種コード': data.code || '',
      '機種名': data.name,
      'シリーズ': data.series || '',
      'ステータス': data.status || '現行',
      '説明': data.description || '',
      'メモ': data.notes || ''
    });
    return { status: 'ok', id, message: '機種を更新しました' };
  }
  appendToSheet('pcb_models', [
    id, data.code || '', data.name, data.series || '',
    data.status || '現行', data.description || '', data.notes || '', getCurrentTime()
  ]);
  return { status: 'ok', id, message: '機種を登録しました' };
}

function deleteModel(id) {
  const boards = getSheetData('pcb_boards').filter(b => b['機種ID'] === id);
  boards.forEach(b => deleteBoard(b['基板ID']));
  deleteRowData('pcb_models', id);
  return '機種と関連データを削除しました';
}

// --------------------------------------------------
// CRUD: 基板
// --------------------------------------------------
function saveBoard(data) {
  setupPCBSheets();
  const id = data.id || generateId();
  if (data.id) {
    _updatePCBRow('pcb_boards', id, {
      '機種ID': data.modelId || '',
      '基板コード': data.code || '',
      '基板名': data.name,
      '改版番号': data.revision || 'Rev.1',
      '改版理由': data.revisionReason || '',
      '説明': data.description || '',
      'メモ': data.notes || ''
    });
    return { status: 'ok', id, message: '基板を更新しました' };
  }
  appendToSheet('pcb_boards', [
    id, data.modelId || '', data.code || '', data.name,
    data.revision || 'Rev.1', data.revisionReason || '',
    data.description || '', data.notes || '', getCurrentTime()
  ]);
  return { status: 'ok', id, message: '基板を登録しました' };
}

function deleteBoard(id) {
  const configs = getSheetData('pcb_configs').filter(c => c['基板ID'] === id);
  configs.forEach(c => deleteConfig(c['構成ID']));
  deleteRowData('pcb_boards', id);
  return '基板と関連データを削除しました';
}

// --------------------------------------------------
// CRUD: 基板構成
// --------------------------------------------------
function saveConfig(data) {
  setupPCBSheets();
  const id = data.id || generateId();
  if (data.id) {
    _updatePCBRow('pcb_configs', id, {
      '基板ID': data.boardId || '',
      '構成名': data.name,
      '構成バージョン': data.version || 'v1.0',
      '層数': data.layers || '',
      '材質': data.material || '',
      'メモ': data.notes || ''
    });
    return { status: 'ok', id, message: '基板構成を更新しました' };
  }
  appendToSheet('pcb_configs', [
    id, data.boardId || '', data.name,
    data.version || 'v1.0', data.layers || '', data.material || '',
    data.notes || '', getCurrentTime()
  ]);
  return { status: 'ok', id, message: '基板構成を登録しました' };
}

function deleteConfig(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('pcb_bom');
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === String(id)) sheet.deleteRow(i + 1);
    }
  }
  deleteRowData('pcb_configs', id);
  return '基板構成と部品表を削除しました';
}

// --------------------------------------------------
// CRUD: 部品 (BOM)
// --------------------------------------------------
function savePart(data) {
  setupPCBSheets();
  const id = data.id || generateId();
  if (data.id) {
    _updatePCBRow('pcb_bom', id, {
      '構成ID': data.configId || '',
      '基板ID': data.boardId || '',
      '参照番号': data.refNo || '',
      '部品番号': data.partNumber || '',
      '部品名': data.partName || '',
      'メーカー': data.manufacturer || '',
      '仕様': data.spec || '',
      '数量': data.quantity || 1,
      '単位': data.unit || '個',
      'リードタイム': data.leadTime || '',
      'メモ': data.notes || ''
    });
    return { status: 'ok', id, message: '部品を更新しました' };
  }
  appendToSheet('pcb_bom', [
    id, data.configId || '', data.boardId || '',
    data.refNo || '', data.partNumber || '', data.partName || '',
    data.manufacturer || '', data.spec || '',
    data.quantity || 1, data.unit || '個',
    data.leadTime || '', data.notes || ''
  ]);
  return { status: 'ok', id, message: '部品を登録しました' };
}

function deletePart(id) {
  deleteRowData('pcb_bom', id);
  return '部品を削除しました';
}

// --------------------------------------------------
// 階層ツリーデータ取得 (フロントエンド用)
// --------------------------------------------------
function getPCBTree() {
  setupPCBSheets();
  const models  = getSheetData('pcb_models');
  const boards  = getSheetData('pcb_boards');
  const configs = getSheetData('pcb_configs');
  const parts   = getSheetData('pcb_bom');

  return models.map(model => {
    const modelId = model['機種ID'];
    const modelBoards = boards
      .filter(b => b['機種ID'] === modelId)
      .map(board => {
        const boardId = board['基板ID'];
        const boardConfigs = configs
          .filter(c => c['基板ID'] === boardId)
          .map(cfg => {
            const configId = cfg['構成ID'];
            const configParts = parts.filter(p => String(p['構成ID']) === String(configId));
            return { ...cfg, parts: configParts, partCount: configParts.length };
          });
        const boardParts = parts.filter(p => String(p['基板ID']) === String(boardId) && (!p['構成ID'] || p['構成ID'] === ''));
        return { ...board, configs: boardConfigs, parts: boardParts, configCount: boardConfigs.length, totalParts: boardParts.length };
      });
    return { ...model, boards: modelBoards, boardCount: modelBoards.length };
  });
}

// --------------------------------------------------
// PCBデータをAIクエリ用テキストに変換
// --------------------------------------------------
function getPCBContext(query) {
  const tree = getPCBTree();
  if (!tree.length) return '';
  // クエリに関連する機種/基板を絞り込む
  const q = query.toLowerCase();
  const relevant = tree.filter(m =>
    JSON.stringify(m).toLowerCase().includes(q.substring(0, 20))
  ).slice(0, 3);
  const src = (relevant.length ? relevant : tree.slice(0, 2));
  return '【基板ナレッジ】\n' + src.map(m => {
    const bds = m.boards.map(b => {
      const cfgs = b.configs.map(c => `      [構成] ${c['構成名']} ${c['構成バージョン']} 部品${c.partCount}点`).join('\n');
      return `  [基板] ${b['基板名']} ${b['改版番号']} (改版理由: ${b['改版理由']})\n${cfgs}`;
    }).join('\n');
    return `[機種] ${m['機種名']} (${m['機種コード']}) ステータス:${m['ステータス']}\n${bds}`;
  }).join('\n\n');
}

// --------------------------------------------------
// 内部ユーティリティ
// --------------------------------------------------
function _updatePCBRow(sheetName, id, fieldMap) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      for (const [col, val] of Object.entries(fieldMap)) {
        const idx = headers.indexOf(col);
        if (idx !== -1) sheet.getRange(i + 1, idx + 1).setValue(val);
      }
      return;
    }
  }
}
