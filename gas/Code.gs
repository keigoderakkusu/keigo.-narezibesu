// ============================================================
// Code.gs — 営業ナレッジベース Pro v2.0
// スクリプトプロパティ: GEMINI_API_KEY（必須）
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('📋 営業ナレッジ Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────
// CORE HELPERS
// ─────────────────────────────────────────────

function _ss() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  const ss = SpreadsheetApp.create('営業ナレッジDB');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function _sheet(name, headers) {
  const ss = _ss();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('white');
    sh.setFrozenRows(1);
  }
  return sh;
}

function _callGemini(prompt, jsonMode) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { error: 'スクリプトプロパティに GEMINI_API_KEY を設定してください' };
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.7,
      ...(jsonMode ? { responseMimeType: 'application/json' } : {})
    }
  };
  try {
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + key,
      { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    if (json.error) return { error: json.error.message };
    const text = json.candidates?.[0]?.content?.parts?.[0]?.text || '';
    if (jsonMode) {
      const cleaned = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      try { return { ok: true, data: JSON.parse(cleaned) }; }
      catch (e) { return { error: 'AI応答のパース失敗。再試行してください。' }; }
    }
    return { ok: true, text };
  } catch (e) {
    return { error: e.message };
  }
}

function _fmt(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), 'Asia/Tokyo', 'MM/dd HH:mm'); } catch (e) { return ''; }
}
function _fmtDate(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), 'Asia/Tokyo', 'yyyy/MM/dd'); } catch (e) { return ''; }
}
function _uuid() { return Utilities.getUuid().split('-')[0]; }
function getCurrentTime() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'); }

function getDbUrl() { return _ss().getUrl(); }

// ─────────────────────────────────────────────
// MEMO
// ─────────────────────────────────────────────

function getMemos(query) {
  const sh = _sheet('memos', ['id', 'title', 'content', 'tags', 'created', 'updated']);
  const rows = sh.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map(r => ({ id: r[0], title: r[1], content: r[2], tags: r[3], created: _fmt(r[4]), updated: _fmt(r[5]) }))
    .reverse();
  if (!query) return rows;
  const q = String(query).toLowerCase();
  return rows.filter(r => (r.title + r.content + r.tags).toLowerCase().includes(q));
}

function saveMemo(memo) {
  const sh = _sheet('memos', ['id', 'title', 'content', 'tags', 'created', 'updated']);
  const now = new Date();
  if (memo.id) {
    const all = sh.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(memo.id)) {
        sh.getRange(i + 1, 2, 1, 5).setValues([[memo.title || '無題', memo.content || '', memo.tags || '', all[i][4], now]]);
        return { ok: true, id: memo.id };
      }
    }
  }
  const id = _uuid();
  sh.appendRow([id, memo.title || '無題', memo.content || '', memo.tags || '', now, now]);
  return { ok: true, id };
}

function deleteMemo(id) {
  const sh = _sheet('memos', ['id', 'title', 'content', 'tags', 'created', 'updated']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

// ─────────────────────────────────────────────
// TASKS + AI PRIORITY ANALYSIS
// ─────────────────────────────────────────────

function getTasks() {
  const sh = _sheet('tasks', ['id', 'title', 'deadline', 'category', 'status', 'created']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  if (rows.length === 0) {
    // サンプルタスクを挿入
    const today = new Date();
    const tomorrow = new Date(today); tomorrow.setDate(today.getDate() + 1);
    const nextWeek = new Date(today); nextWeek.setDate(today.getDate() + 5);
    const samples = [
      [_uuid(), '〇〇社への見積書作成・送付', today, '見積', '未完了', today],
      [_uuid(), '月次営業報告書の提出', tomorrow, '内部作業', '未完了', today],
      [_uuid(), '△△社フォローアップメール送信', today, '営業', '未完了', today],
      [_uuid(), '担当製品カタログの確認・メモ更新', nextWeek, '製品知識', '未完了', today],
    ];
    const sh2 = _sheet('tasks', ['id', 'title', 'deadline', 'category', 'status', 'created']);
    samples.forEach(s => sh2.appendRow(s));
    return getTasks();
  }
  return rows.map(r => ({
    id: r[0], title: r[1],
    deadline: r[2] ? _fmtDate(r[2]) : '',
    category: r[3], status: r[4], created: _fmt(r[5])
  })).reverse();
}

function saveTask(task) {
  const sh = _sheet('tasks', ['id', 'title', 'deadline', 'category', 'status', 'created']);
  const now = new Date();
  const deadline = task.deadline ? new Date(task.deadline) : '';
  if (task.id) {
    const all = sh.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(task.id)) {
        sh.getRange(i + 1, 2, 1, 4).setValues([[task.title, deadline, task.category || '通常', task.status || all[i][4]]]);
        return { ok: true };
      }
    }
  }
  const id = _uuid();
  sh.appendRow([id, task.title, deadline, task.category || '通常', '未完了', now]);
  return { ok: true, id };
}

function toggleTaskStatus(id) {
  const sh = _sheet('tasks', ['id', 'title', 'deadline', 'category', 'status', 'created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) {
      const next = all[i][4] === '完了' ? '未完了' : '完了';
      sh.getRange(i + 1, 5).setValue(next);
      return { ok: true, status: next };
    }
  }
  return { ok: false };
}

function deleteTask(id) {
  const sh = _sheet('tasks', ['id', 'title', 'deadline', 'category', 'status', 'created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

function analyzeTasks() {
  const tasks = getTasks().filter(t => t.status !== '完了');
  if (!tasks.length) return { error: '未完了のタスクがありません。タスクを追加してから分析してください。' };
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd(E)');
  const list = tasks.map((t, i) =>
    `${i + 1}. ${t.title}（期限:${t.deadline || '未設定'} / 種別:${t.category}）`
  ).join('\n');
  const prompt = `あなたは営業職のタスク管理アドバイザーです。
今日は${today}です。以下の未完了タスクを分析し、優先順位とアドバイスをJSONで返してください。

${list}

JSON形式（必ずこの形式で返すこと）:
{
  "situation": "今日の状況を励ます一言（30字以内）",
  "tasks": [
    {
      "rank": 1,
      "originalIndex": 1,
      "priority": "高",
      "reason": "この優先度にした理由（40字以内）",
      "tip": "効率よくこなすコツ（50字以内）",
      "time": "推定作業時間（例: 30分）"
    }
  ],
  "advice": "今日全体へのアドバイス（100字以内）"
}`;
  const result = _callGemini(prompt, true);
  if (!result.ok) return result;
  if (result.data.tasks) {
    result.data.tasks = result.data.tasks.map(t => ({
      ...t,
      taskId: tasks[t.originalIndex - 1]?.id,
      title: tasks[t.originalIndex - 1]?.title || '',
      deadline: tasks[t.originalIndex - 1]?.deadline || '',
      category: tasks[t.originalIndex - 1]?.category || '',
    }));
  }
  return result;
}

// ─────────────────────────────────────────────
// EMAIL TEMPLATES
// ─────────────────────────────────────────────

function getTemplates() {
  const sh = _sheet('templates', ['id', 'name', 'subject', 'body', 'created']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  if (rows.length === 0) {
    const defs = [
      ['見積書送付', '【見積書】{顧客名}様_{商品名}_{日付}',
        '{顧客名} 様\n\nいつもお世話になっております。{自分の名前}です。\nご依頼いただきました見積書をお送りいたします。\nご確認のほどよろしくお願いいたします。\n\nご不明な点はお気軽にご連絡ください。\n\n{自分の名前}'],
      ['訪問後お礼', '【御礼】本日のご面談_{顧客名}様',
        '{顧客名} 様\n\n本日はお忙しい中、貴重なお時間をいただきありがとうございました。\n{議題}についてご説明の機会をいただき大変勉強になりました。\n引き続きどうぞよろしくお願いいたします。\n\n{自分の名前}'],
      ['資料確認フォロー', 'ご確認のお願い_{顧客名}様',
        '{顧客名} 様\n\nいつもお世話になっております。{自分の名前}です。\n先日お送りした{資料名}のご確認状況はいかがでしょうか。\nご不明な点がございましたらお気軽にご連絡ください。\n\nよろしくお願いいたします。'],
      ['納期回答', '【納期回答】{商品名}_{顧客名}様',
        '{顧客名} 様\n\nいつもお世話になっております。{自分の名前}です。\nお問い合わせいただいた{商品名}の納期についてご回答いたします。\n\n納期：{納期}\n数量：{数量}\n\nご発注のほどよろしくお願いいたします。'],
      ['欠品・遅延連絡', '【重要】{商品名} 納期遅延のご連絡_{顧客名}様',
        '{顧客名} 様\n\nいつもお世話になっております。{自分の名前}です。\nご注文いただいております{商品名}について、誠に申し訳ございませんが\n納期が{変更後納期}に変更となりました。\nご迷惑をおかけしますことを深くお詫び申し上げます。\n何卒ご理解いただけますようお願い申し上げます。'],
    ];
    defs.forEach(d => sh.appendRow([_uuid(), d[0], d[1], d[2], new Date()]));
    return getTemplates();
  }
  return rows.map(r => ({ id: r[0], name: r[1], subject: r[2], body: r[3] }));
}

function saveTemplate(tpl) {
  const sh = _sheet('templates', ['id', 'name', 'subject', 'body', 'created']);
  if (tpl.id) {
    const all = sh.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(tpl.id)) {
        sh.getRange(i + 1, 2, 1, 3).setValues([[tpl.name, tpl.subject, tpl.body]]);
        return { ok: true };
      }
    }
  }
  const id = _uuid();
  sh.appendRow([id, tpl.name, tpl.subject, tpl.body, new Date()]);
  return { ok: true, id };
}

function deleteTemplate(id) {
  const sh = _sheet('templates', ['id', 'name', 'subject', 'body', 'created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

function createGmailDraft(to, subject, body) {
  GmailApp.createDraft(to || '', subject, body);
  return { ok: true };
}

// ─────────────────────────────────────────────
// EMAIL PROOFREADING (AI添削)
// ─────────────────────────────────────────────

function proofreadEmail(emailText, context) {
  const prompt = `あなたは日本のビジネスメールのプロフェッショナルです。
営業担当者として送るメール文を添削してください。

${context ? '【補足情報】' + context + '\n\n' : ''}===メール文===
${emailText}
===============

以下のJSON形式で返してください:
{
  "score": 85,
  "verdict": "総合評価の一言",
  "issues": [
    {"type": "敬語", "text": "具体的な問題点と改善案"}
  ],
  "good_points": ["良かった点"],
  "corrected_body": "修正後の本文全文（敬語・表現を最適化）"
}`;
  return _callGemini(prompt, true);
}

// ─────────────────────────────────────────────
// OCR / PDF TEXT EXTRACTION
// ─────────────────────────────────────────────

function processFile(base64, mimeType, prompt) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { error: 'GEMINI_API_KEY が設定されていません' };
  const instruction = prompt ||
    'この画像またはPDFのテキストをすべて正確に抽出してください。表は表形式を維持し、余分なコメントは不要です。';
  const payload = {
    contents: [{ parts: [{ text: instruction }, { inline_data: { mime_type: mimeType, data: base64 } }] }],
    generationConfig: { temperature: 0 }
  };
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + key,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
  );
  const json = JSON.parse(res.getContentText());
  if (json.error) return { error: json.error.message };
  return { ok: true, text: json.candidates?.[0]?.content?.parts?.[0]?.text || '' };
}

// ─────────────────────────────────────────────
// DAILY LOG（業務日誌）
// ─────────────────────────────────────────────

function getDailyLogs(limit) {
  const sh = _sheet('daily_logs', ['id', 'date', 'work', 'achievements', 'issues', 'hours', 'tags', 'created']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({
      id: r[0], date: _fmtDate(r[1]), work: r[2],
      achievements: r[3], issues: r[4], hours: r[5], tags: r[6], created: _fmt(r[7])
    })).reverse();
  return limit ? rows.slice(0, limit) : rows;
}

function saveDailyLog(log) {
  const sh = _sheet('daily_logs', ['id', 'date', 'work', 'achievements', 'issues', 'hours', 'tags', 'created']);
  const now = new Date();
  const date = log.date ? new Date(log.date) : now;
  if (log.id) {
    const all = sh.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(log.id)) {
        sh.getRange(i + 1, 2, 1, 7).setValues([[date, log.work || '', log.achievements || '', log.issues || '', Number(log.hours) || 0, log.tags || '', now]]);
        return { ok: true };
      }
    }
  }
  const id = _uuid();
  sh.appendRow([id, date, log.work || '', log.achievements || '', log.issues || '', Number(log.hours) || 0, log.tags || '', now]);
  return { ok: true, id };
}

function deleteDailyLog(id) {
  const sh = _sheet('daily_logs', ['id', 'date', 'work', 'achievements', 'issues', 'hours', 'tags', 'created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

// ─────────────────────────────────────────────
// HR EVALUATION SUPPORT（人事評価サポート）
// ─────────────────────────────────────────────

function getHrGoals() {
  const sh = _sheet('hr_goals', ['id', 'term', 'title', 'kpi', 'detail', 'created']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  if (rows.length === 0) {
    // 営業2年目に合ったサンプル目標を挿入
    const term = '2025年度上期（4〜9月）';
    const samples = [
      ['既存顧客リピート受注率の向上', 'リピート受注件数 前期比110%以上（月次確認）',
        '既存顧客への定期フォロー頻度を月1回以上に設定し、課題や要望を早期にキャッチして先回りの提案活動を行う。商談記録をナレッジベースに蓄積し引き継ぎ品質も向上させる。'],
      ['業務効率化（見積・メール作成時間50%削減）', '見積書1件あたりの作成時間 現状比50%削減',
        'メールテンプレートとナレッジベースを整備し、過去の見積書や事例をすぐに参照できる仕組みを構築する。定型業務の手順書化も行い、ミスの削減と時間短縮を同時に実現する。'],
      ['製品・技術知識の習得', '技術的質問の自己解決率80%以上を目指す',
        '担当製品（基板・部品・機種構成）の基礎知識を習得し、顧客からの基本的な技術的質問に自力で回答できるようにする。週1回の自己学習時間を確保し、学んだ内容はメモとして蓄積する。'],
      ['新規顧客接点の創出', '新規顧客への提案活動 月2件以上',
        '既存顧客の紹介や展示会等を活用し、新規顧客との接点を月2件以上作る。提案内容をテンプレート化して再現性を高める。'],
    ];
    samples.forEach(s => sh.appendRow([_uuid(), term, s[0], s[1], s[2], new Date()]));
    return getHrGoals();
  }
  return rows.map(r => ({ id: r[0], term: r[1], title: r[2], kpi: r[3], detail: r[4] }));
}

function saveHrGoal(goal) {
  const sh = _sheet('hr_goals', ['id', 'term', 'title', 'kpi', 'detail', 'created']);
  if (goal.id) {
    const all = sh.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(goal.id)) {
        sh.getRange(i + 1, 2, 1, 4).setValues([[goal.term, goal.title, goal.kpi || '', goal.detail || '']]);
        return { ok: true };
      }
    }
  }
  const id = _uuid();
  sh.appendRow([id, goal.term, goal.title, goal.kpi || '', goal.detail || '', new Date()]);
  return { ok: true, id };
}

function deleteHrGoal(id) {
  const sh = _sheet('hr_goals', ['id', 'term', 'title', 'kpi', 'detail', 'created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

function generateEvaluation(term) {
  const goals = getHrGoals();
  const logs = getDailyLogs(60);
  if (!goals.length) return { error: '目標が設定されていません。先に「目標管理」タブで目標を設定してください。' };
  const goalText = goals.map((g, i) =>
    `【目標${i + 1}】${g.title}\nKPI: ${g.kpi}\n詳細: ${g.detail}`
  ).join('\n\n');
  const logText = logs.length
    ? logs.slice(0, 30).map(l =>
        `[${l.date}] ${l.work.slice(0, 60)} ／ 成果:${l.achievements.slice(0, 40)}`
      ).join('\n')
    : '（業務日誌の記録がまだありません。日誌を記録すると評価の精度が上がります）';
  const prompt = `あなたは人事評価のプロフェッショナルです。
以下の情報をもとに、製造業向け営業担当者（現職2年目）の人事評価用自己評価文を作成してください。

【評価期間】${term || '直近期'}
【設定目標】
${goalText}

【業務日誌（抜粋）】
${logText}

以下のJSON形式で返してください:
{
  "achievements": [
    "達成・実績の箇条書き（3〜5項目、具体的に）"
  ],
  "self_eval": "自己評価本文（280字程度、謙虚かつ積極的なトーンで目標への取り組みを具体的に記述）",
  "appeal_points": [
    "特にアピールすべきポイント（2〜3項目）"
  ],
  "growth": "成長・改善の記録（120字程度）",
  "next_goals": [
    "次期に向けた目標案（2〜3項目）"
  ],
  "interview_tips": [
    "評価面談でのアドバイス（2〜3項目）"
  ]
}`;
  return _callGemini(prompt, true);
}

function generateGoalAdvice(situation) {
  const prompt = `あなたは人事評価目標設定の専門家です。
製造業向け営業担当者（部品・基板・機種構成等を扱う、現職2年目）が
人事評価で良い評価を得るためのSMART目標を提案してください。

【現在の状況・課題】
${situation}

以下のJSON形式で返してください:
{
  "goals": [
    {
      "category": "カテゴリ（売上/効率化/スキル/顧客関係/チーム等）",
      "title": "目標タイトル（簡潔に）",
      "detail": "具体的な目標内容（測定可能な指標を含む、120字程度）",
      "kpi": "測定指標（例: 〇〇件以上、〇〇%向上）",
      "why": "この目標を設定する効果・理由（60字程度）"
    }
  ],
  "advice": "目標設定全体へのアドバイス（120字程度）",
  "caution": "よくある失敗と注意点（60字程度）"
}`;
  return _callGemini(prompt, true);
}

// ─────────────────────────────────────────────
// AI CHAT（営業専門アシスタント）
// ─────────────────────────────────────────────

function aiChat(messages) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { error: 'GEMINI_API_KEY が設定されていません' };
  const systemPrompt = `あなたは製造業向け営業職のプロフェッショナルアシスタントです。
対象者は部品・基板・機種構成を扱う営業担当で現職2年目です。
業務効率化、顧客対応、見積作成、社内調整、製品知識習得、スキルアップについてアドバイスします。
【重要】回答は必ず日本語で。簡潔・具体的・実践的に。長文は避け箇条書きを活用してください。`;
  const contents = (messages || []).map(m => ({
    role: m.role === 'user' ? 'user' : 'model',
    parts: [{ text: m.content }]
  }));
  if (!contents.length) return { error: 'メッセージがありません' };
  const payload = {
    contents,
    systemInstruction: { parts: [{ text: systemPrompt }] },
    generationConfig: { temperature: 0.8, maxOutputTokens: 1024 }
  };
  try {
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + key,
      { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    if (json.error) return { error: json.error.message };
    return { ok: true, text: json.candidates?.[0]?.content?.parts?.[0]?.text || '' };
  } catch (e) {
    return { error: e.message };
  }
}

// ─────────────────────────────────────────────
// CHECKLIST
// ─────────────────────────────────────────────

// ─────────────────────────────────────────────
// CUSTOMERS（顧客カルテ）
// ─────────────────────────────────────────────

function getCustomers(query) {
  const sh = _sheet('customers', ['id','company','name','role','phone','email','rank','last_contact','notes','tags','created']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({
      id: r[0], company: r[1], name: r[2], role: r[3],
      phone: r[4], email: r[5], rank: r[6],
      last_contact: r[7] ? _fmtDate(r[7]) : '',
      notes: r[8], tags: r[9], created: _fmt(r[10])
    }))
    .sort((a, b) => ({ 'A': 0, 'B': 1, 'C': 2 }[a.rank] - ({ 'A': 0, 'B': 1, 'C': 2 }[b.rank] || 2) || 0));
  if (!query) return rows;
  const q = String(query).toLowerCase();
  return rows.filter(r => (r.company + r.name + r.tags + r.notes).toLowerCase().includes(q));
}

function saveCustomer(c) {
  const sh = _sheet('customers', ['id','company','name','role','phone','email','rank','last_contact','notes','tags','created']);
  const lc = c.last_contact ? new Date(c.last_contact) : '';
  if (c.id) {
    const all = sh.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(c.id)) {
        sh.getRange(i + 1, 2, 1, 9).setValues([[c.company || '', c.name || '', c.role || '', c.phone || '', c.email || '', c.rank || 'B', lc, c.notes || '', c.tags || '']]);
        return { ok: true };
      }
    }
  }
  const id = _uuid();
  sh.appendRow([id, c.company || '', c.name || '', c.role || '', c.phone || '', c.email || '', c.rank || 'B', lc, c.notes || '', c.tags || '', new Date()]);
  return { ok: true, id };
}

function deleteCustomer(id) {
  const sh = _sheet('customers', ['id','company','name','role','phone','email','rank','last_contact','notes','tags','created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

// ─────────────────────────────────────────────
// MEETINGS（商談メモ）
// ─────────────────────────────────────────────

function getMeetings(limit) {
  const sh = _sheet('meetings', ['id','date','company','person','purpose','content','my_actions','their_actions','next_date','status','created']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({
      id: r[0], date: _fmtDate(r[1]), company: r[2], person: r[3],
      purpose: r[4], content: r[5], my_actions: r[6], their_actions: r[7],
      next_date: r[8] ? _fmtDate(r[8]) : '', status: r[9]
    })).reverse();
  return limit ? rows.slice(0, limit) : rows;
}

function saveMeeting(m) {
  const sh = _sheet('meetings', ['id','date','company','person','purpose','content','my_actions','their_actions','next_date','status','created']);
  const now = new Date();
  const date = m.date ? new Date(m.date) : now;
  const nextDate = m.next_date ? new Date(m.next_date) : '';
  if (m.id) {
    const all = sh.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][0]) === String(m.id)) {
        sh.getRange(i + 1, 2, 1, 9).setValues([[date, m.company || '', m.person || '', m.purpose || '', m.content || '', m.my_actions || '', m.their_actions || '', nextDate, m.status || 'フォロー中']]);
        return { ok: true };
      }
    }
  }
  const id = _uuid();
  sh.appendRow([id, date, m.company || '', m.person || '', m.purpose || '', m.content || '', m.my_actions || '', m.their_actions || '', nextDate, m.status || 'フォロー中', now]);
  return { ok: true, id };
}

function deleteMeeting(id) {
  const sh = _sheet('meetings', ['id','date','company','person','purpose','content','my_actions','their_actions','next_date','status','created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(id)) { sh.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

function generateFollowupEmail(meetingId) {
  const m = getMeetings(null).find(x => x.id === meetingId);
  if (!m) return { error: '商談記録が見つかりません' };
  const prompt = `あなたは製造業向け営業担当者です。以下の商談記録をもとに訪問後フォローアップメールを作成してください。

顧客: ${m.company} ${m.person}様 / 日付: ${m.date}
商談目的: ${m.purpose}
商談内容: ${m.content}
自分の宿題: ${m.my_actions || 'なし'}
先方の宿題: ${m.their_actions || 'なし'}
次回予定: ${m.next_date || '未定'}

要件: 件名(30字以内)と本文(300字以内)。自分の宿題は対応予定を明記。先方の宿題は丁寧に確認。次回への橋渡しで締める。`;
  return _callGemini(prompt, false);
}

function generateMeetingPrep(company, purpose, pastContent) {
  const prompt = `製造業向け営業のプロとして、以下の商談前の準備ブリーフィングをJSONで作成してください。

顧客: ${company} / 目的: ${purpose}
過去の商談: ${pastContent || '（初回訪問または記録なし）'}

JSON形式:
{
  "opening": "最初の一言・アイスブレイク（具体的に）",
  "key_points": ["必ず押さえるポイント（3つ）"],
  "questions": ["顧客に聞くべき質問（3つ）"],
  "watch_out": ["気をつけるべきリスク・注意点"],
  "goal": "この商談のゴール（一文）"
}`;
  return _callGemini(prompt, true);
}

// ─────────────────────────────────────────────
// MONTHLY NUMBERS（数値目標）
// ─────────────────────────────────────────────

function getNumbers(yearMonth) {
  const sh = _sheet('monthly_numbers', ['id','year_month','category','target','actual','unit','created']);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({ id: r[0], year_month: r[1], category: r[2], target: Number(r[3]) || 0, actual: Number(r[4]) || 0, unit: r[5] }));
  return yearMonth ? rows.filter(r => r.year_month === yearMonth) : rows;
}

function saveNumber(n) {
  const sh = _sheet('monthly_numbers', ['id','year_month','category','target','actual','unit','created']);
  const all = sh.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (String(all[i][0]) === String(n.id) || (all[i][1] === n.year_month && all[i][2] === n.category)) {
      sh.getRange(i + 1, 2, 1, 5).setValues([[n.year_month, n.category, Number(n.target) || 0, Number(n.actual) || 0, n.unit || '']]);
      return { ok: true, id: all[i][0] };
    }
  }
  const id = _uuid();
  sh.appendRow([id, n.year_month, n.category, Number(n.target) || 0, Number(n.actual) || 0, n.unit || '', new Date()]);
  return { ok: true, id };
}

function analyzeNumbers(yearMonth) {
  const nums = getNumbers(yearMonth);
  if (!nums.length) return { error: 'この月のデータがありません' };
  const numText = nums.map(n => {
    const rate = n.target > 0 ? Math.round(n.actual / n.target * 100) : 0;
    return `${n.category}: 目標${n.target}${n.unit} 実績${n.actual}${n.unit} 達成率${rate}%`;
  }).join('\n');
  const prompt = `営業担当者の${yearMonth}の実績データです。
${numText}

以下のJSON形式でコメントをください:
{
  "overall": "全体評価（20字以内、必ず励ます）",
  "best": "最もできている点（具体的に）",
  "focus": "今注力すべきこと（具体的に）",
  "tip": "残り期間で達成率を上げるための具体的なアドバイス（80字以内）"
}`;
  return _callGemini(prompt, true);
}

// ─────────────────────────────────────────────
// WEEKLY REPORT（週次報告書AI生成）
// ─────────────────────────────────────────────

function generateWeeklyReport() {
  const logs = getDailyLogs(7);
  const doneTasks = getTasks().filter(t => t.status === '完了').slice(0, 10);
  if (!logs.length) return { error: '業務日誌がありません。日誌を記録してから実行してください。' };
  const logText = logs.map(l => `[${l.date}] ${l.work} / 成果: ${l.achievements}`).join('\n');
  const taskText = doneTasks.length ? doneTasks.map(t => `・${t.title}`).join('\n') : 'なし';
  const prompt = `以下の業務日誌をもとに、上司への週次業務報告書をJSON形式で作成してください。

【業務日誌（直近）】
${logText}

【完了タスク】
${taskText}

JSON形式:
{
  "subject": "週次報告の件名（簡潔に30字以内）",
  "summary": "今週の業務サマリー（150字以内）",
  "achievements": ["主な成果・完了事項（3〜5項目）"],
  "issues": ["課題・懸念事項（あれば、なければ空配列）"],
  "next_week": ["来週の主な予定・方針（3項目）"],
  "message": "上司へのひと言（30字以内、前向きに）"
}`;
  return _callGemini(prompt, true);
}

// ─────────────────────────────────────────────
// QUOTE CHECKER
// ─────────────────────────────────────────────
// 見積書（Excel/PDF/画像）を4項目でAI自動チェック
// 対応: ①誤字脱字 ②計算ミス ③必須項目 ④商品マスタ照合
// ─────────────────────────────────────────────

/** Driveフォルダを監視して全ファイルを自動チェック（5分おきトリガー用） */
function checkQuotesInFolder() {
  const props = PropertiesService.getScriptProperties();
  const inboxFolderId = props.getProperty('QUOTE_INBOX_FOLDER_ID');
  if (!inboxFolderId) { Logger.log('QUOTE_INBOX_FOLDER_ID 未設定'); return; }

  const files = DriveApp.getFolderById(inboxFolderId).getFiles();
  const processed = [];

  while (files.hasNext()) {
    const file = files.next();
    const mime = file.getMimeType();
    const name = file.getName();
    try {
      let quoteData = null;
      if (mime === MimeType.MICROSOFT_EXCEL ||
          mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
          name.endsWith('.xlsx') || name.endsWith('.xls')) {
        quoteData = readExcelFile(file);
      } else if (mime === MimeType.PDF) {
        quoteData = readWithGeminiOCR(file, 'application/pdf');
      } else if (mime === MimeType.JPEG || mime === MimeType.PNG ||
                 mime === 'image/jpeg' || mime === 'image/png') {
        quoteData = readWithGeminiOCR(file, mime);
      } else {
        moveQuoteFile(file, 'QUOTE_ERROR_FOLDER_ID');
        continue;
      }
      const results = runAllChecks(quoteData, file.getName());
      logCheckResults(file.getName(), file.getUrl(), results);
      sendNotificationEmail(file.getName(), file.getUrl(), results);
      saveQuoteToMemos(quoteData, results, file);
      moveQuoteFile(file, 'QUOTE_DONE_FOLDER_ID');
      processed.push(name);
    } catch(e) {
      Logger.log('エラー [' + name + ']: ' + e.message);
      moveQuoteFile(file, 'QUOTE_ERROR_FOLDER_ID');
    }
  }
  return '処理完了: ' + processed.length + '件';
}

/** Excel→GスプレッドシートへOCR読み込み（Drive API 有効化が必要） */
function readExcelFile(file) {
  const tempResource = { title: '__tmp_quote_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS };
  const tempFile = Drive.Files.insert(tempResource, file.getBlob());
  try {
    const ss = SpreadsheetApp.openById(tempFile.id);
    return parseExcelData(ss.getActiveSheet().getDataRange().getValues(), file.getName());
  } finally {
    try { Drive.Files.remove(tempFile.id); } catch(e) {}
  }
}

function parseExcelData(data, fileName) {
  const result = {
    fileName, format: 'Excel', rawData: data,
    company:'', quoteNo:'', quoteDate:'', expiryDate:'',
    staffName:'', contactInfo:'', ourCompany:'', deliveryDate:'',
    items:[], subtotal:0, tax:0, total:0,
    rawText: data.map(r => r.join('\t')).join('\n')
  };
  for (let i = 0; i < Math.min(20, data.length); i++) {
    const row = data[i]; const rowStr = row.join(' ');
    if (rowStr.match(/御見積|宛先|会社名/) && !result.company)
      result.company = extractValueNearby(row, ['御見積','宛先','会社名']);
    if (rowStr.match(/見積番号|見積No/))
      result.quoteNo = extractValueNearby(row, ['見積番号','見積No','No.']);
    if (rowStr.match(/見積日|作成日/))
      result.quoteDate = extractValueNearby(row, ['見積日','作成日']);
    if (rowStr.match(/有効期限|期限/))
      result.expiryDate = extractValueNearby(row, ['有効期限','期限']);
    if (rowStr.match(/担当者|担当/))
      result.staffName = extractValueNearby(row, ['担当者','担当']);
    if (rowStr.match(/納期|納入/))
      result.deliveryDate = extractValueNearby(row, ['納期','納入']);
  }
  let inItems = false;
  for (let i = 0; i < data.length; i++) {
    const row = data[i]; const rowStr = row.join(' ');
    if (rowStr.match(/品名|商品名|品目/) && rowStr.match(/数量|個数/) && rowStr.match(/単価|価格/)) {
      inItems = true; continue;
    }
    if (inItems) {
      if (rowStr.match(/小計|合計/) && !rowStr.match(/品名|商品/)) {
        const nums = row.filter(c => typeof c === 'number' || (String(c).match(/^[\d,]+$/) && c));
        if (nums.length > 0) {
          const val = parseFloat(String(nums[nums.length-1]).replace(/,/g,''));
          if (rowStr.match(/税込|合計/)) result.total = val;
          else if (rowStr.match(/消費税|税額/)) result.tax = val;
          else result.subtotal = val;
        }
        continue;
      }
      if (row.every(c => c === '' || c === null || c === undefined)) continue;
      const item = parseItemRow(row);
      if (item) result.items.push(item);
    }
  }
  if (result.subtotal === 0 && result.items.length > 0) {
    result.subtotal = result.items.reduce((s,i) => s + (i.subtotal||0), 0);
    result.tax   = Math.round(result.subtotal * 0.1);
    result.total = result.subtotal + result.tax;
  }
  return result;
}

function extractValueNearby(row, keywords) {
  for (let i = 0; i < row.length; i++) {
    if (keywords.some(k => String(row[i]).includes(k))) {
      for (let j = i+1; j < Math.min(i+4, row.length); j++) {
        if (row[j] !== '' && row[j] !== null) return String(row[j]);
      }
    }
  }
  return '';
}

function parseItemRow(row) {
  const nums=[]; const texts=[];
  row.forEach(c => {
    if (typeof c === 'number') nums.push(c);
    else if (typeof c === 'string' && c.match(/^[\d,]+(\.\d+)?$/) && c.length > 0)
      nums.push(parseFloat(c.replace(/,/g,'')));
    else if (typeof c === 'string' && c.trim()) texts.push(c.trim());
  });
  if (nums.length < 2 && texts.length === 0) return null;
  return { name:texts[0]||'', code:texts[1]||'', qty:nums[0]||0, price:nums[1]||0, subtotal:nums[2]||(nums[0]*nums[1])||0 };
}

/** PDF・画像をGemini OCRで読み込み */
function readWithGeminiOCR(file, mimeType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が未設定です');
  const base64 = Utilities.base64Encode(file.getBlob().getBytes());
  const prompt = `この見積書を読み取りJSONのみ返してください。説明文不要。
{"company":"宛先会社名","quoteNo":"見積番号","quoteDate":"YYYY-MM-DD","expiryDate":"YYYY-MM-DD","staffName":"担当者名","contactInfo":"連絡先","ourCompany":"自社名","deliveryDate":"納期","items":[{"name":"商品名","code":"型番","qty":数量,"price":単価,"subtotal":小計}],"subtotal":税抜合計,"tax":消費税,"total":税込合計,"rawText":"全文テキスト"}`;
  const payload = {
    contents: [{ parts:[{inline_data:{mime_type:mimeType,data:base64}},{text:prompt}] }],
    generationConfig: { temperature:0 }
  };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
  const res = UrlFetchApp.fetch(url, { method:'POST', contentType:'application/json', payload:JSON.stringify(payload), muteHttpExceptions:true });
  const json = JSON.parse(res.getContentText());
  if (!json.candidates?.[0]) throw new Error('Gemini OCRエラー');
  let text = json.candidates[0].content.parts[0].text.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
  try {
    const parsed = JSON.parse(text);
    parsed.fileName = file.getName();
    parsed.format = mimeType.includes('pdf') ? 'PDF' : 'Image';
    return parsed;
  } catch(e) {
    return { fileName:file.getName(), format:'OCR_Error', rawText:text, items:[], company:'', quoteNo:'', quoteDate:'', expiryDate:'', staffName:'', contactInfo:'', ourCompany:'', deliveryDate:'', subtotal:0, tax:0, total:0 };
  }
}

/** 4つのチェックを一括実行 */
function runAllChecks(quoteData, fileName) {
  const results = {
    fileName,
    timestamp: getCurrentTime(),
    check1_typo:     checkTypoAndVariation(quoteData),
    check2_calc:     checkCalculation(quoteData),
    check3_required: checkRequiredFields(quoteData),
    check4_product:  checkProductMaster(quoteData),
    overallStatus:   'OK'
  };
  const checks = [results.check1_typo, results.check2_calc, results.check3_required, results.check4_product];
  results.overallStatus = checks.some(c=>c.status==='ERROR') ? 'ERROR' : checks.some(c=>c.status==='WARNING') ? 'WARNING' : 'OK';
  return results;
}

/** チェック①：誤字脱字・表記ゆれ（Gemini） */
function checkTypoAndVariation(quoteData) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { status:'SKIP', issues:[], message:'APIキー未設定' };
  const textToCheck = [quoteData.rawText||'', (quoteData.items||[]).map(i=>i.name+' '+i.code).join('\n')].join('\n').substring(0,8000);
  const prompt = `以下の見積書テキストの誤字脱字・表記ゆれを検出し、JSONのみ返してください。
{"issues":[{"type":"誤字脱字|表記ゆれ|その他","found":"問題箇所","suggestion":"修正案"}],"hasIssue":true/false}
テキスト:\n${textToCheck}`;
  try {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    const res = UrlFetchApp.fetch(url, { method:'POST', contentType:'application/json',
      payload: JSON.stringify({ contents:[{parts:[{text:prompt}]}], generationConfig:{temperature:0} }),
      muteHttpExceptions:true });
    let text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
    text = text.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
    const parsed = JSON.parse(text);
    return { status: parsed.hasIssue?'WARNING':'OK', issues:parsed.issues||[], message: parsed.hasIssue?`${parsed.issues.length}件の表記ゆれを検出`:'問題なし' };
  } catch(e) { return { status:'SKIP', issues:[], message:'チェックエラー: '+e.message }; }
}

/** チェック②：計算ミス（GASロジック） */
function checkCalculation(quoteData) {
  const TOLERANCE = 1;
  if (!quoteData.items || quoteData.items.length === 0)
    return { status:'WARNING', issues:[], message:'明細行を取得できませんでした（フォーマット確認が必要）' };
  const issues = [];
  let calcSubtotal = 0;
  quoteData.items.forEach((item, idx) => {
    const expected = item.qty * item.price;
    const actual   = item.subtotal || 0;
    if (Math.abs(expected - actual) > TOLERANCE && expected > 0) {
      issues.push({ type:'小計ミス', found:`${idx+1}行目「${item.name}」: ${item.qty}×${item.price.toLocaleString()}=${expected.toLocaleString()} → 記載:${actual.toLocaleString()}`, suggestion:`小計は${expected.toLocaleString()}円が正しい` });
    }
    calcSubtotal += (actual > 0 ? actual : expected);
  });
  if (quoteData.subtotal > 0 && Math.abs(quoteData.subtotal - calcSubtotal) > TOLERANCE)
    issues.push({ type:'税抜合計ミス', found:`明細合計:${calcSubtotal.toLocaleString()} → 記載:${quoteData.subtotal.toLocaleString()}`, suggestion:`税抜合計は${calcSubtotal.toLocaleString()}円が正しい` });
  const base = quoteData.subtotal > 0 ? quoteData.subtotal : calcSubtotal;
  const expectedTax = Math.round(base * 0.1);
  if (quoteData.tax > 0 && Math.abs(quoteData.tax - expectedTax) > TOLERANCE)
    issues.push({ type:'消費税ミス', found:`記載:${quoteData.tax.toLocaleString()} → 計算値:${expectedTax.toLocaleString()}`, suggestion:`消費税は${expectedTax.toLocaleString()}円（10%）` });
  const expectedTotal = base + (quoteData.tax > 0 ? quoteData.tax : expectedTax);
  if (quoteData.total > 0 && Math.abs(quoteData.total - expectedTotal) > TOLERANCE)
    issues.push({ type:'税込合計ミス', found:`記載:${quoteData.total.toLocaleString()} → 計算値:${expectedTotal.toLocaleString()}`, suggestion:`税込合計は${expectedTotal.toLocaleString()}円が正しい` });
  return { status:issues.length>0?'ERROR':'OK', issues, message:issues.length>0?`${issues.length}件の計算ミスを検出`:`計算OK（税抜:${calcSubtotal.toLocaleString()}円 / 税込:${expectedTotal.toLocaleString()}円）` };
}

/** チェック③：必須項目チェック（GASロジック） */
function checkRequiredFields(quoteData) {
  const required = [
    {key:'company',label:'宛先会社名'},{key:'quoteNo',label:'見積番号'},
    {key:'quoteDate',label:'見積日'},{key:'expiryDate',label:'有効期限'},
    {key:'ourCompany',label:'自社名'},{key:'staffName',label:'担当者名'},
    {key:'deliveryDate',label:'納期'}
  ];
  const issues = [];
  required.forEach(f => {
    const val = quoteData[f.key];
    if (!val || String(val).trim()==='' || String(val)==='undefined')
      issues.push({ type:'記入漏れ', found:`「${f.label}」が空欄`, suggestion:`${f.label}を記入してください` });
  });
  if (!quoteData.items || quoteData.items.length===0)
    issues.push({ type:'記入漏れ', found:'明細行が1件もありません', suggestion:'商品・サービスの明細を記入してください' });
  return { status:issues.length>0?'ERROR':'OK', issues, message:issues.length>0?`${issues.length}件の必須項目が未記入`:'全必須項目OK' };
}

/** チェック④：商品マスタ照合（スプレッドシートDB） */
function checkProductMaster(quoteData) {
  const masterSheetId = PropertiesService.getScriptProperties().getProperty('PRODUCT_MASTER_SHEET_ID');
  if (!masterSheetId) return { status:'SKIP', issues:[], message:'PRODUCT_MASTER_SHEET_ID 未設定のためスキップ' };
  if (!quoteData.items || quoteData.items.length===0) return { status:'OK', issues:[], message:'明細なし' };
  try {
    const masterData = SpreadsheetApp.openById(masterSheetId).getSheets()[0].getDataRange().getValues();
    if (masterData.length < 2) return { status:'SKIP', issues:[], message:'商品マスタにデータがありません' };
    const headers = masterData[0].map(h=>String(h).trim());
    const codeCol  = headers.findIndex(h=>h.match(/型番|品番|コード/));
    const priceCol = headers.findIndex(h=>h.match(/価格|単価/));
    const statusCol= headers.findIndex(h=>h.match(/ステータス|廃番|状態/));
    const masterMap = {};
    masterData.slice(1).forEach(row => {
      const code = codeCol>=0 ? String(row[codeCol]).trim().toLowerCase() : '';
      if (code) masterMap[code] = { price:priceCol>=0?row[priceCol]:null, status:statusCol>=0?String(row[statusCol]):'' };
    });
    const issues = [];
    quoteData.items.forEach((item, idx) => {
      if (!item.code || !item.code.trim()) {
        if (item.name) issues.push({ type:'型番未記入', found:`${idx+1}行目「${item.name}」の型番が空欄`, suggestion:'型番を記入してください' });
        return;
      }
      const m = masterMap[item.code.trim().toLowerCase()];
      if (!m) {
        issues.push({ type:'不明な型番', found:`${idx+1}行目「${item.code}」がマスタに存在しません`, suggestion:'型番を確認するかマスタに登録してください' });
      } else {
        if (m.status.match(/廃番|廃止|discontinued/i))
          issues.push({ type:'廃番商品', found:`${idx+1}行目「${item.code}」は廃番です`, suggestion:'後継機種への変更を検討してください' });
        if (m.price && item.price > 0) {
          const mp = parseFloat(String(m.price).replace(/[^0-9.]/g,''));
          const diff = Math.abs(item.price-mp)/mp;
          if (diff > 0.2) issues.push({ type:'価格乖離', found:`${idx+1}行目「${item.code}」単価:${item.price.toLocaleString()}円（マスタ:${mp.toLocaleString()}円, 乖離${Math.round(diff*100)}%）`, suggestion:'単価を確認してください' });
        }
      }
    });
    return { status:issues.length>0?'WARNING':'OK', issues, message:issues.length>0?`${issues.length}件のマスタ不一致`:`全${quoteData.items.length}件の型番照合OK` };
  } catch(e) { return { status:'SKIP', issues:[], message:'マスタ照合エラー: '+e.message }; }
}

/** チェック結果をスプレッドシートにログ保存 */
function logCheckResults(fileName, fileUrl, results) {
  const sheet = _sheet('quote_check_logs', ['日時','ファイル名','ファイルURL','総合ステータス','①誤字脱字','②計算','③必須項目','④型番照合','問題詳細']);
  const SE = { OK:'✅', WARNING:'⚠️', ERROR:'❌', SKIP:'⏭️' };
  const allIssues = [
    ...results.check1_typo.issues, ...results.check2_calc.issues,
    ...results.check3_required.issues, ...results.check4_product.issues
  ].map(i=>`[${i.type}] ${i.found}`).join(' | ');
  sheet.appendRow([
    results.timestamp, fileName, fileUrl,
    SE[results.overallStatus]+' '+results.overallStatus,
    SE[results.check1_typo.status]+' '+results.check1_typo.message,
    SE[results.check2_calc.status]+' '+results.check2_calc.message,
    SE[results.check3_required.status]+' '+results.check3_required.message,
    SE[results.check4_product.status]+' '+results.check4_product.message,
    allIssues || 'なし'
  ]);
  const last = sheet.getLastRow();
  if (results.overallStatus==='ERROR')   sheet.getRange(last,1,1,9).setBackground('#fee2e2');
  else if (results.overallStatus==='WARNING') sheet.getRange(last,1,1,9).setBackground('#fef9c3');
}

/** ERRORまたはWARNINGのみメール通知 */
function sendNotificationEmail(fileName, fileUrl, results) {
  if (results.overallStatus==='OK') return;
  const email = PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
  if (!email) return;
  const SE = { OK:'✅', WARNING:'⚠️', ERROR:'❌', SKIP:'⏭️' };
  const subject = `【見積書チェック】${SE[results.overallStatus]} ${results.overallStatus} - ${fileName}`;
  const lines = [
    `ファイル: ${fileName}`, `URL: ${fileUrl}`, `総合: ${results.overallStatus}`, `日時: ${results.timestamp}`, '─────',
    `① 誤字脱字: ${SE[results.check1_typo.status]} ${results.check1_typo.message}`,
    `② 計算    : ${SE[results.check2_calc.status]} ${results.check2_calc.message}`,
    `③ 必須項目: ${SE[results.check3_required.status]} ${results.check3_required.message}`,
    `④ 型番照合: ${SE[results.check4_product.status]} ${results.check4_product.message}`,
  ];
  [results.check1_typo,results.check2_calc,results.check3_required,results.check4_product].forEach(c => {
    (c.issues||[]).forEach(i => lines.push(`   → [${i.type}] ${i.found}  修正案: ${i.suggestion||'-'}`));
  });
  GmailApp.sendEmail(email, subject, lines.join('\n'));
}

/** チェック済み見積書データをメモとして保存 */
function saveQuoteToMemos(quoteData, results, file) {
  try {
    const content = [
      `【見積書チェック済み】${quoteData.fileName}`,
      `宛先: ${quoteData.company}  見積No: ${quoteData.quoteNo}  見積日: ${quoteData.quoteDate}`,
      `税込合計: ${(quoteData.total||0).toLocaleString()}円  有効期限: ${quoteData.expiryDate}  担当: ${quoteData.staffName}`,
      `総合チェック: ${results.overallStatus}`,
      `─── 明細 ───`,
      (quoteData.items||[]).map(i=>`${i.name} ${i.code} × ${i.qty} = ${(i.subtotal||0).toLocaleString()}円`).join('\n'),
      `DriveURL: ${file.getUrl()}`
    ].join('\n');
    saveMemo({ id:_uuid(), title:'見積書: '+quoteData.fileName, content, tags:'見積書,チェック済み', updatedAt:getCurrentTime() });
  } catch(e) { Logger.log('メモ保存エラー: '+e.message); }
}

function moveQuoteFile(file, destPropKey) {
  const id = PropertiesService.getScriptProperties().getProperty(destPropKey);
  if (!id) return;
  try {
    const dest = DriveApp.getFolderById(id);
    dest.addFile(file);
    const parents = file.getParents();
    while (parents.hasNext()) parents.next().removeFile(file);
  } catch(e) { Logger.log('移動エラー: '+e.message); }
}

/** UIからファイルIDを指定して単一チェック（メインUI呼び出し口） */
function checkSingleFileById(fileId) {
  const file = DriveApp.getFileById(fileId);
  const mime = file.getMimeType();
  const name = file.getName();
  let quoteData;
  if (mime === MimeType.MICROSOFT_EXCEL ||
      mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      name.endsWith('.xlsx') || name.endsWith('.xls')) {
    quoteData = readExcelFile(file);
  } else if (mime === MimeType.PDF) {
    quoteData = readWithGeminiOCR(file, 'application/pdf');
  } else if (mime === MimeType.JPEG || mime === MimeType.PNG) {
    quoteData = readWithGeminiOCR(file, mime);
  } else {
    return { error: true, message: '対応外フォーマット: ' + mime };
  }
  const results = runAllChecks(quoteData, name);
  logCheckResults(name, file.getUrl(), results);
  sendNotificationEmail(name, file.getUrl(), results);
  return results;
}

/** ログ履歴をUI向けに返す */
function getQuoteCheckLogs(limit) {
  limit = limit || 50;
  const sheet = _ss().getSheetByName('quote_check_logs');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).reverse().slice(0, limit).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    return obj;
  });
}

/** 初期セットアップ（初回のみGASエディタから手動実行） */
function setupQuoteChecker() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction()==='checkQuotesInFolder')) {
    ScriptApp.newTrigger('checkQuotesInFolder').timeBased().everyMinutes(5).create();
  }
  return ['✅ 初期セットアップ完了','スクリプトプロパティに以下を設定してください:',
    'QUOTE_INBOX_FOLDER_ID / QUOTE_DONE_FOLDER_ID / QUOTE_ERROR_FOLDER_ID',
    'NOTIFY_EMAIL（任意）/ PRODUCT_MASTER_SHEET_ID（任意）'].join('\n');
}

// ─────────────────────────────────────────────
// SELF ANALYSIS
// ─────────────────────────────────────────────

function getSelfAnalysis(period, targetPeriod) {
  // period: 'month' | 'year'
  // targetPeriod: 'YYYY-MM' or 'YYYY'
  let startStr, endStr, label;
  if (period === 'month') {
    const [y, m] = targetPeriod.split('-').map(Number);
    const start = new Date(y, m - 1, 1);
    const end   = new Date(y, m, 0);
    startStr = _fmt(start); endStr = _fmt(end);
    label = y + '年' + m + '月';
  } else {
    const y = parseInt(targetPeriod);
    startStr = y + '-01-01'; endStr = y + '-12-31';
    label = y + '年';
  }

  // ── collect daily logs ──
  const logRows = _sheet('daily_logs', ['id','date','content','mood','tags'])
    .getDataRange().getValues().slice(1)
    .filter(r => r[0] && String(r[1]) >= startStr && String(r[1]) <= endStr);
  const logs = logRows.map(r => ({ date: r[1], content: String(r[2]).substring(0, 120), mood: r[3] }));

  // ── collect done tasks ──
  const taskRows = _sheet('tasks', ['id','title','due','priority','status','createdAt'])
    .getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[4] === 'done');
  const tasks = taskRows.map(r => ({ title: r[1], due: String(r[2]), priority: r[3] }))
    .filter(t => t.due >= startStr && t.due <= endStr);

  // ── collect meetings ──
  const mtgRows = _sheet('meetings', ['id','date','company','contact','purpose','content','followup','nextAction','createdAt'])
    .getDataRange().getValues().slice(1)
    .filter(r => r[0] && String(r[1]) >= startStr && String(r[1]) <= endStr);
  const meetings = mtgRows.map(r => ({ date: r[1], company: r[2], purpose: r[4], nextAction: r[7] }));

  // ── collect numbers ──
  const numRows = _sheet('monthly_numbers', ['id','yearMonth','category','target','actual','unit','updatedAt'])
    .getDataRange().getValues().slice(1).filter(r => r[0]);
  const nums = numRows.filter(r => {
    const ym = String(r[1]);
    return period === 'month' ? ym === targetPeriod : ym.startsWith(targetPeriod);
  }).map(r => ({ ym: r[1], cat: r[2], tgt: Number(r[3]), act: Number(r[4]), unit: r[5] }));

  // ── collect HR goals ──
  const goalRows = _sheet('hr_goals', ['id','term','category','goal','measure','progress','selfEval','createdAt'])
    .getDataRange().getValues().slice(1).filter(r => r[0]);
  const goals = goalRows.map(r => ({ term: r[1], cat: r[2], goal: r[3], progress: r[5], selfEval: r[6] }));

  // ── stats ──
  const moodVals = logs.map(l => parseInt(l.mood)).filter(v => !isNaN(v) && v > 0);
  const avgMood  = moodVals.length ? (moodVals.reduce((a,b) => a+b,0)/moodVals.length).toFixed(1) : '-';
  const companies = [...new Set(meetings.map(m => m.company).filter(Boolean))];
  const numAchieved = nums.filter(n => n.tgt > 0 && n.act >= n.tgt).length;

  // ── build text blocks ──
  const logText  = logs.slice(-40).map(l => `[${l.date}] ${l.content}`).join('\n') || '記録なし';
  const taskText = tasks.slice(-40).map(t => `[${t.priority||'中'}] ${t.title}`).join('\n') || '記録なし';
  const mtgText  = meetings.map(m => `[${m.date}] ${m.company}: ${m.purpose}`).join('\n') || '記録なし';
  const numText  = nums.map(n => {
    const pct = n.tgt > 0 ? Math.round(n.act / n.tgt * 100) : '-';
    return `${n.ym} ${n.cat}: 目標${n.tgt}${n.unit} 実績${n.act}${n.unit} 達成率${pct}%`;
  }).join('\n') || '記録なし';
  const goalText = goals.map(g => `[${g.term}][${g.cat}] ${g.goal} 進捗:${g.progress||'未'} 自己評価:${g.selfEval||'未'}`).join('\n') || '記録なし';

  const prompt = `あなたは製造業（部品・基板・機械設備）の2年目営業担当者のキャリアコーチ兼人事評価アドバイザーです。
以下のデータをもとに ${label} の振り返り自己分析レポートを作成してください。

【集計】業務日誌:${logs.length}件 完了タスク:${tasks.length}件 商談:${meetings.length}件 訪問先:${companies.length}社 平均コンディション:${avgMood}/5 数値目標達成:${numAchieved}/${nums.length}項目

【業務日誌（抜粋）】
${logText}

【完了タスク】
${taskText}

【商談記録】
${mtgText}

【数値実績】
${numText}

【人事評価目標】
${goalText}

以下のJSON形式で分析レポートを生成してください。データが少ない項目は一般的な2年目製造業営業の知見で補完してください：

{
  "period": "${label}",
  "highlights": ["成果・実績ハイライト（3〜5件、具体的数字があれば含める）"],
  "strengths": ["この期間に発揮できた強み・成長した点（3件）"],
  "challenges": ["正直な改善課題（2〜3件、成長機会として前向きに表現）"],
  "goalAchievement": "目標達成度の総評（2〜3文、データに基づき客観的に）",
  "growthScore": 成長度スコア（1〜10の数値のみ）,
  "growthComment": "スコアの根拠（1〜2文）",
  "conditionTrend": "コンディション傾向と分析（1〜2文）",
  "hrSelfPR": "人事評価提出用の自己PR文（200〜300文字、です・ます調、このまま使えるレベルで。具体的な数字・行動・成果を含める）",
  "hrGoalProposal": "次期目標の提案文（150文字前後、SMART形式、このまま提出できる文章）",
  "coachMessage": "コーチからの総合メッセージ（3〜4文、励ましと具体的指摘を含む）",
  "nextActions": ["来期に向けた具体的アクション（3件、すぐ実行できるレベルで）"],
  "hiddenStrength": "データから読み取れる本人が気づいていないかもしれない潜在的強み（1件）"
}`;

  let data;
  try {
    data = JSON.parse(_callGemini(prompt, true));
  } catch(e) {
    data = { error: '分析に失敗しました。再試行してください。period=' + period };
  }

  data.rawStats = {
    logCount: logs.length, taskCount: tasks.length,
    meetingCount: meetings.length, companyCount: companies.length,
    avgMood, numAchieved, numTotal: nums.length
  };
  data.generatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  return data;
}

// ─────────────────────────────────────────────
// NEWS SUMMARY
// ─────────────────────────────────────────────

function getNewsSummary(forceRefresh) {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('NEWS_CACHE');
  if (cached && !forceRefresh) {
    const c = JSON.parse(cached);
    if (Date.now() - c.ts < 6 * 60 * 60 * 1000) return c.data;
  }

  const articles = [];
  const feeds = [
    { url: 'https://www3.nhk.or.jp/rss/news/cat4.xml', label: '経済' },
    { url: 'https://www3.nhk.or.jp/rss/news/cat6.xml', label: '国際' },
    { url: 'https://news.google.com/rss/search?q=%E8%A3%BD%E9%80%A0%E6%A5%AD+%E6%9D%90%E6%96%99%E8%B2%BB+%E9%83%A8%E5%93%81+%E9%AB%98%E9%A8%B0&hl=ja&gl=JP&ceid=JP:ja', label: '材料・部品' },
    { url: 'https://news.google.com/rss/search?q=%E4%B8%AD%E6%9D%B1+%E5%8E%9F%E6%B2%B9+%E3%82%A8%E3%83%8D%E3%83%AB%E3%82%AE%E3%83%BC%E4%BE%A1%E6%A0%BC&hl=ja&gl=JP&ceid=JP:ja', label: '中東・エネルギー' },
    { url: 'https://news.google.com/rss/search?q=%E5%86%86%E7%9B%B8%E5%A0%B4+%E7%82%BA%E6%9B%BF+%E8%BC%B8%E5%87%BA%E5%85%A5&hl=ja&gl=JP&ceid=JP:ja', label: '為替' },
    { url: 'https://news.google.com/rss/search?q=%E5%8D%8A%E5%B0%8E%E4%BD%93+%E9%89%84%E9%8B%BC+%E3%82%A2%E3%83%AB%E3%83%9F+%E4%BE%9B%E7%B5%A6%E9%80%BC%E8%BF%AB&hl=ja&gl=JP&ceid=JP:ja', label: '素材市況' },
  ];

  feeds.forEach(feed => {
    try {
      const res = UrlFetchApp.fetch(feed.url, { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() !== 200) return;
      const xml = XmlService.parse(res.getContentText('UTF-8'));
      const channel = xml.getRootElement().getChild('channel');
      if (!channel) return;
      channel.getChildren('item').slice(0, 7).forEach(item => {
        const title = item.getChildText('title') || '';
        const desc = (item.getChildText('description') || '').replace(/<[^>]+>/g, '').substring(0, 100);
        if (title) articles.push({ category: feed.label, title, desc });
      });
    } catch(e) { Logger.log('RSS error [' + feed.label + ']: ' + e); }
  });

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日');
  const headlines = articles.slice(0, 35).map((a, i) =>
    `${i + 1}. [${a.category}] ${a.title}${a.desc ? ' / ' + a.desc : ''}`
  ).join('\n') || '（RSS取得不可。AIの学習知識に基づき分析します）';

  const prompt = `あなたは製造業（部品・基板・機械設備）の営業担当者を支援するビジネスアナリストです。
本日 ${today} のニュース見出し一覧です（RSS取得）：

${headlines}

以下のJSON形式で要約・分析してください。製造業の2年目営業担当者が朝3分で読める内容にしてください。

{
  "sections": [
    {
      "icon": "絵文字1つ",
      "title": "セクション名（15文字以内）",
      "summary": "3〜4文の要約。具体的な数字・企業名・変動幅があれば含める",
      "salesTip": "このニュースを踏まえた営業トーク・対応アクションの提案（1〜2文）"
    }
  ],
  "briefing": "今日の営業で意識すべき3ポイント（①②③番号付き、改行区切り）",
  "keywords": ["キーワード1","キーワード2","キーワード3","キーワード4","キーワード5"],
  "alertLevel": "green または yellow または red",
  "alertReason": "警戒度の理由（20文字以内）"
}

sectionsは必ず以下4テーマで作成：
1. 材料・部品・エネルギー価格（鉄鋼・アルミ・樹脂・半導体・原油・天然ガス）
2. 中東・地政学リスク（紛争・制裁・ホルムズ海峡・サプライチェーン影響）
3. 為替・金融動向（円ドル・日銀・米FRB・輸出入コストへの影響）
4. 国内製造業トレンド（設備投資・受注動向・景況感・技術動向）`;

  let data;
  try {
    data = JSON.parse(_callGemini(prompt, true));
  } catch(e) {
    data = {
      sections: [{ icon: '⚠️', title: '分析エラー', summary: '「今すぐ再取得」を押してください。', salesTip: '' }],
      briefing: '', keywords: [], alertLevel: 'green', alertReason: ''
    };
  }

  data.articleCount = articles.length;
  data.updatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'M/d HH:mm');
  props.setProperty('NEWS_CACHE', JSON.stringify({ ts: Date.now(), data }));
  return data;
}

// ─────────────────────────────────────────────
// CHECKLIST
// ─────────────────────────────────────────────

function getChecklist(type) {
  const lists = {
    email: ['宛先（To/CC）は正しいか', '件名に顧客名・日付が入っているか', '添付ファイルを忘れていないか',
      '金額・型番を声に出して再確認したか', '誤字脱字はないか', '敬称・敬語は適切か', '署名は正しく入っているか'],
    estimate: ['顧客名・宛名は正しいか', '品番・型番は正確か', '単価・数量・合計を計算し直したか',
      '有効期限は記載したか', '承認者の確認は済んでいるか', 'ファイル名に顧客名と日付を入れたか',
      '前回見積書と差異がないか確認したか'],
    visit: ['訪問先の住所・経路を確認したか', '持参する資料・サンプルを準備したか', '名刺は十分にあるか',
      '前回の打ち合わせ内容を確認したか', 'アポイントの時間・担当者名を確認したか',
      '商談目的・ゴールを決めたか', '相手企業の最新情報を確認したか'],
  };
  return lists[type] || [];
}
