// ============================================================
// Feature_Vector.gs  ── ⑪ ベクトル検索（意味的類似度検索）
// Gemini Embedding API を使って全ナレッジにベクトルを付与し
// キーワード一致ではなく「意味」で検索する
// ============================================================

const PROP_EMBED_MODEL = 'text-embedding-004'; // Gemini Embedding モデル名

// ── シート拡張（setupDatabase に追記相当） ──────────────────
function setupVectorDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = {
    // ベクトルインデックス: 各ソースのEmbeddingを保存
    'vector_index': [
      'ベクトルID','ソース種別(kb/meeting/note/product)',
      'ソースID','テキスト要約(256文字)','埋め込みベクトルJSON',
      '登録日時','最終更新日時'
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
  return 'ベクトルDBの初期化完了';
}

// ── テキスト → Embedding ベクトル取得 ───────────────────────
function getEmbedding(text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP_GEMINI);
  if (!apiKey) throw new Error('GEMINI_API_KEY が未設定です');

  // 256文字に切り詰め（トークン節約 & レート制限回避）
  const truncated = text.replace(/\s+/g, ' ').substring(0, 2000);

  const url = `https://generativelanguage.googleapis.com/v1beta/models/text-embedding-004:embedContent?key=${apiKey}`;
  const payload = {
    model: 'models/text-embedding-004',
    content: { parts: [{ text: truncated }] },
    taskType: 'RETRIEVAL_DOCUMENT'
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(res.getContentText());
  if (!json.embedding) throw new Error('Embedding取得失敗: ' + JSON.stringify(json));
  return json.embedding.values; // float配列（768次元）
}

// ── 全ナレッジを一括インデックス化 ─────────────────────────
// GASエディタから手動実行、または定期トリガーで実行
function buildVectorIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = getVectorIndex(); // 既存のソースID一覧
  const existingIds = new Set(existing.map(r => r['ソースID']));

  let indexed = 0, skipped = 0, errors = 0;
  const results = [];

  // ── KB記事 ──
  const kbArts = rows('kb_articles').filter(a => a['ステータス(公開/下書き)'] !== '下書き');
  kbArts.forEach(a => {
    if (existingIds.has(a['記事ID'])) { skipped++; return; }
    try {
      const text = `${a['タイトル']} ${a['カテゴリ']} ${a['タグ']} ${(a['本文']||'').substring(0,500)}`;
      const vec  = getEmbedding(text);
      results.push([uid(), 'kb', a['記事ID'], text.substring(0,256), JSON.stringify(vec), now(), now()]);
      indexed++;
      Utilities.sleep(500); // レート制限対策
    } catch(e) { errors++; Logger.log('KB索引エラー: ' + e.message); }
  });

  // ── 商談議事録 ──
  const meetings = rows('meetings');
  meetings.forEach(m => {
    if (existingIds.has(m['ID'])) { skipped++; return; }
    try {
      const text = `${m['顧客名']} ${m['機種名']} ${m['基板名']} ${m['内容サマリ']} ${(m['議事録全文']||'').substring(0,300)}`;
      const vec  = getEmbedding(text);
      results.push([uid(), 'meeting', m['ID'], text.substring(0,256), JSON.stringify(vec), now(), now()]);
      indexed++;
      Utilities.sleep(500);
    } catch(e) { errors++; }
  });

  // ── ナレッジメモ ──
  const notes = rows('notes');
  notes.forEach(n => {
    if (existingIds.has(n['ID'])) { skipped++; return; }
    try {
      const text = `${n['タグ']} ${n['機種名']} ${n['基板名']} ${(n['メモ内容']||'').substring(0,400)}`;
      const vec  = getEmbedding(text);
      results.push([uid(), 'note', n['ID'], text.substring(0,256), JSON.stringify(vec), now(), now()]);
      indexed++;
      Utilities.sleep(500);
    } catch(e) { errors++; }
  });

  // ── 製品マスタ ──
  const products = rows('products');
  products.forEach(p => {
    if (existingIds.has(p['ID'])) { skipped++; return; }
    try {
      const text = `${p['製品名']} ${p['ステータス']} ${p['競合優位性']} ${p['改廃理由']}`;
      const vec  = getEmbedding(text);
      results.push([uid(), 'product', p['ID'], text.substring(0,256), JSON.stringify(vec), now(), now()]);
      indexed++;
      Utilities.sleep(500);
    } catch(e) { errors++; }
  });

  // バッチ書き込み
  if (results.length > 0) {
    const sheet = ss.getSheetByName('vector_index');
    results.forEach(r => sheet.appendRow(r));
  }

  return {
    message: `インデックス完了: 新規${indexed}件 / スキップ${skipped}件 / エラー${errors}件`,
    indexed, skipped, errors
  };
}

// ── ベクトルインデックス取得 ────────────────────────────────
function getVectorIndex() {
  return rows('vector_index');
}

// ── コサイン類似度計算 ──────────────────────────────────────
function cosineSimilarity(a, b) {
  let dot = 0, normA = 0, normB = 0;
  for (let i = 0; i < a.length; i++) {
    dot   += a[i] * b[i];
    normA += a[i] * a[i];
    normB += b[i] * b[i];
  }
  return dot / (Math.sqrt(normA) * Math.sqrt(normB));
}

// ── メイン: 意味的ベクトル検索 ─────────────────────────────
/**
 * クエリテキストをベクトル化し、全インデックスとコサイン類似度を計算。
 * 上位 topK 件の実データを返す。
 * @param {string} query    - 検索クエリ
 * @param {number} topK     - 返す件数（デフォルト10）
 * @param {string} srcType  - 絞込み種別（'kb'|'meeting'|'note'|'product'|'all'）
 * @returns {Array} - スコア付き検索結果
 */
function semanticSearch(query, topK, srcType) {
  topK    = topK    || 8;
  srcType = srcType || 'all';

  try {
    // クエリのベクトル化
    const queryVec = getEmbedding(query);

    // インデックス読み込み
    let index = getVectorIndex();
    if (srcType !== 'all') index = index.filter(r => r['ソース種別(kb/meeting/note/product)'] === srcType);

    if (!index.length) return { status: 'empty', results: [], message: 'インデックスが空です。先に buildVectorIndex() を実行してください。' };

    // 類似度計算
    const scored = index.map(row => {
      try {
        const vec   = JSON.parse(row['埋め込みベクトルJSON']);
        const score = cosineSimilarity(queryVec, vec);
        return { ...row, score };
      } catch(e) {
        return { ...row, score: 0 };
      }
    });

    // スコア降順ソート → 上位 topK
    scored.sort((a, b) => b.score - a.score);
    const topResults = scored.slice(0, topK);

    // 実データを取得してエンリッチ
    const enriched = topResults.map(r => {
      const type = r['ソース種別(kb/meeting/note/product)'];
      const sid  = r['ソースID'];
      let detail = {};

      if (type === 'kb') {
        const art = rows('kb_articles').find(a => String(a['記事ID']) === String(sid));
        if (art) detail = { title: art['タイトル'], excerpt: (art['本文']||'').substring(0,200), category: art['カテゴリ'], tags: art['タグ'], id: sid };
      } else if (type === 'meeting') {
        const m = rows('meetings').find(x => String(x['ID']) === String(sid));
        if (m) detail = { title: `【商談】${m['顧客名']} - ${m['内容サマリ']}`, excerpt: m['議事録全文']?.substring(0,200)||'', category: '商談議事録', tags: m['機種名'], id: sid };
      } else if (type === 'note') {
        const n = rows('notes').find(x => String(x['ID']) === String(sid));
        if (n) detail = { title: `【メモ】${n['タグ']}`, excerpt: (n['メモ内容']||'').substring(0,200), category: 'ナレッジメモ', tags: n['機種名'], id: sid };
      } else if (type === 'product') {
        const p = rows('products').find(x => String(x['ID']) === String(sid));
        if (p) detail = { title: `【製品】${p['製品名']}`, excerpt: p['競合優位性']?.substring(0,200)||'', category: '製品マスタ', tags: p['ステータス'], id: sid };
      }

      return {
        type,
        score:   Math.round(r.score * 1000) / 1000,
        summary: r['テキスト要約(256文字)'],
        ...detail
      };
    }).filter(r => r.title); // データが存在するもののみ

    return { status: 'success', results: enriched, query };

  } catch(e) {
    return { status: 'error', message: e.message, results: [] };
  }
}

// ── 記事登録時に自動インデックス化（saveKbArticle から呼び出す） ──
function indexSingleDocument(type, sourceId, text) {
  try {
    // 既存インデックスを更新
    const index = getVectorIndex();
    const existing = index.find(r => r['ソースID'] === sourceId);
    const vec = getEmbedding(text);
    const vecJson = JSON.stringify(vec);

    if (existing) {
      // 既存を更新
      updateRowData('vector_index', existing['ベクトルID'], '埋め込みベクトルJSON', vecJson);
      updateRowData('vector_index', existing['ベクトルID'], '最終更新日時', now());
    } else {
      // 新規追加
      push('vector_index', [uid(), type, sourceId, text.substring(0,256), vecJson, now(), now()]);
    }
    return true;
  } catch(e) {
    Logger.log('自動インデックス失敗: ' + e.message);
    return false;
  }
}

// ── インデックス状況確認 ─────────────────────────────────────
function getVectorIndexStats() {
  const index = getVectorIndex();
  const byType = {};
  index.forEach(r => {
    const t = r['ソース種別(kb/meeting/note/product)'];
    byType[t] = (byType[t] || 0) + 1;
  });
  const totalDocs = {
    kb:      rows('kb_articles').filter(a=>a['ステータス(公開/下書き)']!=='下書き').length,
    meeting: rows('meetings').length,
    note:    rows('notes').length,
    product: rows('products').length,
  };
  return {
    indexed: byType,
    total:   totalDocs,
    coverage: Object.fromEntries(
      Object.entries(totalDocs).map(([k,v]) => [k, v > 0 ? Math.round((byType[k]||0)/v*100) : 0])
    )
  };
}
