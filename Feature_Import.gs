// ============================================================
// Feature_Import.gs  ── ⑧ スプレッドシート一括インポート
// 外部Excel/スプレッドシートを解析してDB（製品・議事録・顧客）に振り分け
// Geminiがヘッダー・内容を自動判定してマッピング
// ============================================================

// ── シート拡張 ──────────────────────────────────────────────
function setupImportDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = {
    'import_logs': [
      'インポートID','実行日時','ファイル名','ファイルID',
      'Drive URL','判定種別','取込件数','スキップ件数',
      'エラー件数','ステータス','詳細メモ'
    ],
  };
  for (const [name, headers] of Object.entries(cfg)) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      const r = s.getRange(1,1,1,headers.length);
      r.setValues([headers]).setFontWeight('bold').setBackground('#1e293b').setFontColor('#fff');
      s.setFrozenRows(1);
    }
  }
  return 'インポートDBの初期化完了';
}

// ── メイン: DriveフォルダIDを指定して一括インポート ─────────
/**
 * 指定フォルダ内のxlsx/csv/Google Sheetsを全てスキャンし、
 * Geminiが種別を判定して各DBシートに振り分ける
 * @param {string} folderId - DriveフォルダのID
 * @returns {object} - インポート結果サマリ
 */
function importFromDriveFolder(folderId) {
  if (!folderId) return { status: 'error', message: 'フォルダIDを指定してください' };

  let folder;
  try { folder = DriveApp.getFolderById(folderId); }
  catch(e) { return { status: 'error', message: 'フォルダが見つかりません: ' + e.message }; }

  const files   = folder.getFiles();
  const results = [];
  let totalImported = 0, totalErrors = 0;

  while (files.hasNext()) {
    const file = files.next();
    const mime = file.getMimeType();

    // 対応ファイル形式のみ処理
    const supported = [
      MimeType.GOOGLE_SHEETS,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // xlsx
      MimeType.CSV,
    ];
    if (!supported.includes(mime)) continue;

    try {
      const result = importSingleFile(file);
      results.push(result);
      totalImported += result.imported || 0;
      totalErrors   += result.errors   || 0;

      // インポートログ記録
      push('import_logs', [
        uid(), now(), file.getName(), file.getId(), file.getUrl(),
        result.detectedType || '不明', result.imported || 0,
        result.skipped || 0, result.errors || 0,
        result.status, result.message || ''
      ]);

    } catch(e) {
      results.push({ file: file.getName(), status: 'error', message: e.message });
      totalErrors++;
    }
  }

  return {
    status:  totalErrors === 0 ? 'success' : 'partial',
    message: `インポート完了: ${results.length}ファイル処理 / ${totalImported}件取込 / ${totalErrors}件エラー`,
    results, totalImported, totalErrors
  };
}

// ── 単一ファイルのインポート処理 ────────────────────────────
function importSingleFile(file) {
  const mime = file.getMimeType();
  let ssData;

  // ファイルをスプレッドシートデータとして読み込む
  if (mime === MimeType.GOOGLE_SHEETS) {
    ssData = readGoogleSheets(file.getId());
  } else if (mime === MimeType.CSV) {
    ssData = readCsvFile(file);
  } else {
    // xlsxはGoogleシートに変換してから読む
    ssData = readXlsxAsSheets(file);
  }

  if (!ssData || !ssData.length) {
    return { file: file.getName(), status: 'skip', message: 'データなし', imported: 0 };
  }

  // Geminiで種別判定
  const detectedType = detectDataType(ssData[0]);

  // 種別に応じてインポート
  let result;
  switch (detectedType) {
    case 'products':  result = importAsProducts(ssData[0]);  break;
    case 'meetings':  result = importAsMeetings(ssData[0]);  break;
    case 'notes':     result = importAsNotes(ssData[0]);     break;
    case 'models':    result = importAsModels(ssData[0]);    break;
    case 'customers': result = importAsCustomers(ssData[0]); break;
    default:
      result = importAsNotes(ssData[0]); // 判定不明はメモとして取込
      result.message = '種別不明のためナレッジメモとして取込みました';
  }

  return {
    file: file.getName(),
    url:  file.getUrl(),
    detectedType,
    ...result
  };
}

// ── Google Sheets 読み込み ───────────────────────────────────
function readGoogleSheets(fileId) {
  const ss     = SpreadsheetApp.openById(fileId);
  const sheets = ss.getSheets();
  return sheets.map(sheet => {
    const data = sheet.getDataRange().getValues();
    return { sheetName: sheet.getName(), headers: data[0] || [], rows: data.slice(1) };
  }).filter(s => s.rows.length > 0);
}

// ── CSV 読み込み ─────────────────────────────────────────────
function readCsvFile(file) {
  const text = file.getBlob().getDataAsString('UTF-8');
  const lines = text.split('\n').filter(l => l.trim());
  if (lines.length < 2) return [];
  const headers = parseCsvLine(lines[0]);
  const rows    = lines.slice(1).map(parseCsvLine);
  return [{ sheetName: file.getName(), headers, rows }];
}

function parseCsvLine(line) {
  const result = []; let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') { inQ = !inQ; }
    else if (c === ',' && !inQ) { result.push(cur.trim()); cur = ''; }
    else { cur += c; }
  }
  result.push(cur.trim());
  return result;
}

// ── xlsx → Google Sheets変換して読み込み ─────────────────────
function readXlsxAsSheets(file) {
  try {
    const converted = Drive.Files.insert(
      { title: file.getName() + '_temp', mimeType: MimeType.GOOGLE_SHEETS },
      file.getBlob()
    );
    const result = readGoogleSheets(converted.id);
    Drive.Files.remove(converted.id); // 一時ファイル削除
    return result;
  } catch(e) {
    Logger.log('xlsx変換失敗: ' + e.message);
    return [];
  }
}

// ── Geminiによるデータ種別判定 ──────────────────────────────
function detectDataType(sheetData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP_GEMINI);
  if (!apiKey) return guessTypeByHeaders(sheetData.headers);

  const headerSample = sheetData.headers.join(', ');
  const rowSample    = (sheetData.rows[0] || []).join(', ');

  const prompt = `以下のスプレッドシートのヘッダーとサンプルデータを見て、このデータが何の種別か1単語で答えてください。
選択肢: products（製品・部品・在庫）, meetings（議事録・商談・打合せ）, notes（メモ・ナレッジ・情報）, models（機種・モデル・型番）, customers（顧客・取引先・会社）, other
ヘッダー: ${headerSample}
サンプル: ${rowSample}
回答（1単語のみ）:`;

  try {
    const res = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
      { method:'post', contentType:'application/json', muteHttpExceptions:true,
        payload: JSON.stringify({ contents:[{ parts:[{text:prompt}]}], generationConfig:{temperature:0, maxOutputTokens:10}}) }
    );
    const answer = JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text?.trim().toLowerCase() || 'other';
    const valid  = ['products','meetings','notes','models','customers'];
    return valid.includes(answer) ? answer : guessTypeByHeaders(sheetData.headers);
  } catch(e) {
    return guessTypeByHeaders(sheetData.headers);
  }
}

// ── ヘッダーキーワードによるフォールバック判定 ─────────────
function guessTypeByHeaders(headers) {
  const h = headers.join(' ').toLowerCase();
  if (/製品|部品|型番|price|product|在庫|unit/.test(h))    return 'products';
  if (/議事録|商談|顧客|client|meeting|訪問|打合/.test(h)) return 'meetings';
  if (/機種|model|モデル|型式/.test(h))                    return 'models';
  if (/顧客|取引先|company|customer|会社/.test(h))         return 'customers';
  return 'notes';
}

// ── Geminiによるカラムマッピング生成 ─────────────────────────
/**
 * 外部ファイルのヘッダーをDBカラムに自動マッピング
 * 例: "品名" → "製品名", "Customer" → "顧客名"
 */
function buildColumnMapping(srcHeaders, targetFields, sheetData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP_GEMINI);

  // Gemini API なしのフォールバックマッピング
  const fallback = buildFallbackMapping(srcHeaders, targetFields);
  if (!apiKey) return fallback;

  const prompt = `以下のソースヘッダーを、ターゲットフィールドにマッピングしてください。
対応できないものは null にしてください。JSON形式のみで回答してください。

ソースヘッダー: ${JSON.stringify(srcHeaders)}
ターゲットフィールド: ${JSON.stringify(targetFields)}
サンプルデータ（1行目）: ${JSON.stringify(sheetData.rows[0] || [])}

回答形式: {"ターゲットフィールド名": "ソースヘッダー名 または null", ...}`;

  try {
    const res = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`,
      { method:'post', contentType:'application/json', muteHttpExceptions:true,
        payload: JSON.stringify({ contents:[{parts:[{text:prompt}]}], generationConfig:{temperature:0}}) }
    );
    const text = JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    const clean = text.replace(/```json|```/g,'').trim();
    return JSON.parse(clean);
  } catch(e) {
    return fallback;
  }
}

// ── フォールバックマッピング（キーワード一致） ───────────────
function buildFallbackMapping(srcHeaders, targetFields) {
  const mapping = {};
  const synonyms = {
    '製品名':    ['製品名','品名','製品','型番','item','product','name','品目'],
    '価格':      ['価格','単価','金額','price','cost','¥'],
    'ステータス':['ステータス','状態','status','区分'],
    '顧客名':    ['顧客名','会社名','取引先','顧客','client','customer','company'],
    '機種名':    ['機種名','機種','型式','model','モデル'],
    '基板名':    ['基板名','基板','board','pcb'],
    '内容サマリ':['サマリ','概要','内容','summary','件名','subject','タイトル'],
    '議事録全文':['全文','詳細','内容','議事録','memo','備考'],
    'タグ':      ['タグ','tag','分類','カテゴリ','category'],
    'メモ内容':  ['メモ','内容','コメント','備考','memo','note'],
    '機種ID':    ['機種id','modelid','model_id'],
    '仕様概要':  ['仕様','spec','specification','概要'],
    '担当者':    ['担当者','担当','person','owner','charge'],
    'カテゴリ':  ['カテゴリ','分類','category','区分','種別'],
    '競合優位性':['優位性','強み','競合','advantage'],
  };
  targetFields.forEach(tf => {
    const syns = synonyms[tf] || [tf];
    const found = srcHeaders.find(sh =>
      syns.some(s => sh.toLowerCase().replace(/[\s_-]/g,'').includes(s.toLowerCase().replace(/[\s_-]/g,'')))
    );
    mapping[tf] = found || null;
  });
  return mapping;
}

// ── getValue helper ──────────────────────────────────────────
function getVal(row, headers, colName) {
  if (!colName) return '';
  const idx = headers.indexOf(colName);
  return idx >= 0 ? String(row[idx] || '').trim() : '';
}

// ── 製品マスタとしてインポート ───────────────────────────────
function importAsProducts(sheetData) {
  const targetFields = ['製品名','価格','ステータス','競合優位性','改廃理由','後継機種'];
  const mapping = buildColumnMapping(sheetData.headers, targetFields, sheetData);
  const existing = rows('products').map(r => r['製品名']);
  let imported = 0, skipped = 0, errors = 0;

  sheetData.rows.forEach(row => {
    try {
      const name = getVal(row, sheetData.headers, mapping['製品名']) || row[0];
      if (!name) { skipped++; return; }
      if (existing.includes(name)) { skipped++; return; } // 重複スキップ
      push('products', [
        uid(),
        name,
        getVal(row, sheetData.headers, mapping['価格']),
        getVal(row, sheetData.headers, mapping['ステータス']) || '現行品',
        getVal(row, sheetData.headers, mapping['改廃理由']),
        getVal(row, sheetData.headers, mapping['後継機種']),
        getVal(row, sheetData.headers, mapping['競合優位性']),
      ]);
      imported++;
    } catch(e) { errors++; }
  });

  return { status:'success', imported, skipped, errors, message: `製品マスタに${imported}件取込` };
}

// ── 商談議事録としてインポート ───────────────────────────────
function importAsMeetings(sheetData) {
  const targetFields = ['顧客名','機種名','基板名','内容サマリ','議事録全文'];
  const mapping = buildColumnMapping(sheetData.headers, targetFields, sheetData);
  let imported = 0, skipped = 0, errors = 0;

  sheetData.rows.forEach(row => {
    try {
      const client = getVal(row, sheetData.headers, mapping['顧客名']) || row[0];
      if (!client) { skipped++; return; }
      const id = uid();
      push('meetings', [
        id, now(), client,
        '', getVal(row, sheetData.headers, mapping['機種名']),
        '', getVal(row, sheetData.headers, mapping['基板名']),
        getVal(row, sheetData.headers, mapping['内容サマリ']),
        getVal(row, sheetData.headers, mapping['議事録全文']),
        '', '', ''
      ]);
      imported++;
    } catch(e) { errors++; }
  });

  return { status:'success', imported, skipped, errors, message: `商談議事録に${imported}件取込` };
}

// ── ナレッジメモとしてインポート ─────────────────────────────
function importAsNotes(sheetData) {
  const targetFields = ['タグ','メモ内容','機種名','基板名'];
  const mapping = buildColumnMapping(sheetData.headers, targetFields, sheetData);
  let imported = 0, skipped = 0, errors = 0;

  sheetData.rows.forEach(row => {
    try {
      const content = getVal(row, sheetData.headers, mapping['メモ内容']) || row.join(' | ');
      if (!content || content.length < 3) { skipped++; return; }
      const id = uid();
      push('notes', [
        id, now(),
        getVal(row, sheetData.headers, mapping['タグ']),
        content,
        '', getVal(row, sheetData.headers, mapping['機種名']),
        '', getVal(row, sheetData.headers, mapping['基板名']),
        '', '', ''
      ]);
      imported++;
    } catch(e) { errors++; }
  });

  return { status:'success', imported, skipped, errors, message: `ナレッジメモに${imported}件取込` };
}

// ── 機種マスタとしてインポート ───────────────────────────────
function importAsModels(sheetData) {
  const targetFields = ['機種名','カテゴリ','仕様概要','担当者'];
  const mapping = buildColumnMapping(sheetData.headers, targetFields, sheetData);
  const existing = rows('models').map(r => r['機種名']);
  let imported = 0, skipped = 0, errors = 0;

  sheetData.rows.forEach(row => {
    try {
      const name = getVal(row, sheetData.headers, mapping['機種名']) || row[0];
      if (!name) { skipped++; return; }
      if (existing.includes(name)) { skipped++; return; }
      push('models', [
        uid(), name,
        getVal(row, sheetData.headers, mapping['カテゴリ']),
        getVal(row, sheetData.headers, mapping['仕様概要']),
        getVal(row, sheetData.headers, mapping['担当者']),
        now(), ''
      ]);
      imported++;
    } catch(e) { errors++; }
  });

  return { status:'success', imported, skipped, errors, message: `機種マスタに${imported}件取込` };
}

// ── 顧客データとして議事録マスタにインポート ─────────────────
function importAsCustomers(sheetData) {
  // 顧客マスタは議事録の顧客名列に一意エントリとして登録
  const targetFields = ['顧客名','内容サマリ','機種名'];
  const mapping = buildColumnMapping(sheetData.headers, targetFields, sheetData);
  let imported = 0, skipped = 0;

  sheetData.rows.forEach(row => {
    const client = getVal(row, sheetData.headers, mapping['顧客名']) || row[0];
    if (!client) { skipped++; return; }
    push('meetings', [uid(), now(), client, '', getVal(row, sheetData.headers, mapping['機種名']), '', '', getVal(row, sheetData.headers, mapping['内容サマリ']) || '顧客マスタからインポート', '', '', '', '']);
    imported++;
  });

  return { status:'success', imported, skipped, errors:0, message: `顧客データを議事録に${imported}件取込` };
}

// ── 指定URLのスプレッドシートを直接インポート ─────────────────
function importFromSpreadsheetUrl(url, targetType) {
  try {
    // URLからIDを抽出
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) return { status: 'error', message: 'URLからスプレッドシートIDを取得できませんでした' };
    const ssId = match[1];

    const ssData = readGoogleSheets(ssId);
    if (!ssData.length) return { status: 'error', message: 'データが空です' };

    const sheetData = ssData[0];
    const detectedType = targetType || detectDataType(sheetData);

    let result;
    switch(detectedType) {
      case 'products':  result = importAsProducts(sheetData);  break;
      case 'meetings':  result = importAsMeetings(sheetData);  break;
      case 'notes':     result = importAsNotes(sheetData);     break;
      case 'models':    result = importAsModels(sheetData);    break;
      case 'customers': result = importAsCustomers(sheetData); break;
      default:          result = importAsNotes(sheetData);     break;
    }

    // インポートログ記録
    push('import_logs', [uid(), now(), `URL指定インポート`, ssId, url, detectedType, result.imported||0, result.skipped||0, result.errors||0, result.status, result.message||'']);

    return { ...result, detectedType, url };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}

// ── インポートログ取得 ────────────────────────────────────────
function getImportLogs(limit) {
  const all = rows('import_logs').sort((a,b) => new Date(b['実行日時']) - new Date(a['実行日時']));
  return limit ? all.slice(0, limit) : all;
}

// ── プレビュー（実際には取込まず内容確認のみ） ────────────────
function previewImport(url) {
  try {
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) return { status: 'error', message: 'URLが正しくありません' };
    const ssData = readGoogleSheets(match[1]);
    if (!ssData.length) return { status: 'error', message: 'データが空です' };

    const sheetData    = ssData[0];
    const detectedType = detectDataType(sheetData);
    const sampleRows   = sheetData.rows.slice(0, 5);

    return {
      status: 'success',
      detectedType,
      headers:    sheetData.headers,
      totalRows:  sheetData.rows.length,
      sampleRows,
      sheetName:  sheetData.sheetName,
      message: `${sheetData.rows.length}行を検出。種別:「${detectedType}」として取込予定`
    };
  } catch(e) {
    return { status: 'error', message: e.message };
  }
}
