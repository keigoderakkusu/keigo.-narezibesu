// ==========================================
// Code.gs
// バックエンドロジック・APIルーティング
// ==========================================

const SCRIPT_PROP_GEMINI_KEY = 'GEMINI_API_KEY';

/**
 * Webアプリへのアクセス時の処理
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('営業ナレッジナビゲーター')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 別ファイル(CSS/JavaScript)を動的に読み込むための関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 【DB自動構築ロジック】初回実行時のみ実行
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 要件定義に基づくシートとカラム構造
  const sheetsConfig = {
    'products': ['ID', '製品名', '価格', 'ステータス', '改廃理由', '後継機種', '競合優位性'],
    'meetings': ['ID', '登録日時', '顧客名', '関連基板・機種', '内容サマリ', '議事録全文'],
    'notes': ['ID', '登録日時', 'タグ(Obsidianタグ等)', 'メモ内容'],
    'notebooklm': ['ID', 'ノートブック名', '関連製品', '共有URL', 'Enterprise連携フラグ'],
    'qalogs': ['日時', 'ユーザー入力', 'AI回答内容'],
    'sources': ['ID', 'ファイル名', 'URL', 'タイプ', '連携日時'] // 将来のDrive連携用拡張枠
  };

  for (const sheetName in sheetsConfig) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    const headers = sheetsConfig[sheetName];
    
    // ヘッダーが未設定（1行目が空白）の場合のみ初期化
    if (sheet.getLastRow() === 0) {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1e293b"); // モダンなダークブルー
      headerRange.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
  }
  return "データベースの初期化・構築が完了しました！シートを確認してください。";
}

// ----------------------------------------------------
// DBユーティリティ関数
// ----------------------------------------------------

function generateId() {
  return Utilities.getUuid();
}

function getCurrentTime() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
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
    rows = rows.slice(-limit); // トークン節約のため最新データのみを取得可能にする
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
// フロントエンド用 データ取得系API
// ----------------------------------------------------

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getCount = (name) => {
    const s = ss.getSheetByName(name);
    return s ? Math.max(0, s.getLastRow() - 1) : 0;
  };

  const notebookLMs = getSheetData('notebooklm').reverse(); // 最新順

  return {
    productsCount: getCount('products'),
    meetingsCount: getCount('meetings'),
    notesCount: getCount('notes'),
    notebookLMs: notebookLMs
  };
}

// ----------------------------------------------------
// フロントエンド用 データ登録系API
// ----------------------------------------------------

function saveProduct(data) {
  appendToSheet('products', [
    generateId(), data.name, data.price, data.status, data.reason, data.successor, data.advantage
  ]);
  return "製品情報を正常に登録しました。";
}

function saveMeeting(data) {
  appendToSheet('meetings', [
    generateId(), getCurrentTime(), data.client, data.related, data.summary, data.fulltext
  ]);
  return "商談議事録を正常に登録しました。AIの学習対象に追加されました。";
}

function saveNote(data) {
  appendToSheet('notes', [
    generateId(), getCurrentTime(), data.tags, data.content
  ]);
  return "ナレッジメモ（Obsidian形式等）を正常に登録しました。";
}

function saveNotebookLM(data) {
  appendToSheet('notebooklm', [
    generateId(), data.name, data.related, data.url, '未連携(手動)' // Enterprise連携までのプレースホルダー
  ]);
  return "NotebookLMリンク台帳に登録しました。";
}

// ----------------------------------------------------
// AI ナレッジ検索＆新人ロープレ機能 (Gemini API 連携)
// ----------------------------------------------------

function processAIQuery(query, mode) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_GEMINI_KEY);
  if (!apiKey) {
    return "システムエラー: Gemini APIキーが設定されていません。管理者に連絡するか、初期設定手順を確認してください。";
  }

  // 最新の社内データをコンテキストとして抽出 (モデルのトークン上限・速度を考慮し最新データを取得)
  const products = getSheetData('products', 100); 
  const meetings = getSheetData('meetings', 30);
  const notes = getSheetData('notes', 30);

  // AIに渡す社内コンテキスト文字列の構築
  let contextText = "【製品マスタ情報】\n";
  contextText += products.map(p => `- 製品名: ${p['製品名']} (価格: ${p['価格']}, 状態: ${p['ステータス']})\n  優位性: ${p['競合優位性']}\n  後継機種: ${p['後継機種']}, 改廃理由: ${p['改廃理由']}`).join("\n");
  
  contextText += "\n\n【最新の商談議事録】\n";
  contextText += meetings.map(m => `- [${m['登録日時']}] 顧客: ${m['顧客名']} (関連: ${m['関連基板・機種']})\n  サマリ: ${m['内容サマリ']}`).join("\n");
  
  contextText += "\n\n【社内ナレッジメモ】\n";
  contextText += notes.map(n => `- タグ[${n['タグ(Obsidianタグ等)']}] 内容: ${n['メモ内容']}`).join("\n");

  // アプリケーションモードに応じたプロンプトの切り替え
  let systemInstruction = "";
  if (mode === 'roleplay') {
    systemInstruction = `あなたは製造業企業の厳格な「顧客（購買部長や技術責任者）」です。提供された当社の【製品マスタ情報】や【最新の商談議事録】の背景情報を踏まえつつ、新人営業担当者の提案や質問に対して、あえて鋭い指摘、価格交渉、または代替品の性能に対する現実的な懸念を投げかけてください。応答は対話（ロープレ会話）形式とし、簡潔でリアリティのある口調にしてください。`;
  } else {
    systemInstruction = `あなたは組織の知を統合した優秀な「セールス・イネーブルメントAI」です。ユーザーからの質問に対し、以下の【社内データ】のみを情報源として、最も適切で根拠に基づいた回答を作成してください。社内で解決すべき方針が見える場合はアドバイスも添えてください。ただし、提供された社内データに該当する情報が全くない場合は、推測で語らず「社内データベースには該当情報が見つかりません」と正直に回答してください。\n\n${contextText}`;
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const payload = {
    "system_instruction": {
      "parts": { "text": systemInstruction }
    },
    "contents": [{
      "parts": [{ "text": query }]
    }],
    "generationConfig": {
      // ロープレ時は変動性を持たせ、検索時は一貫性を重視
      "temperature": mode === 'roleplay' ? 0.8 : 0.2 
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    let aiResponse = "";
    if (json.candidates && json.candidates.length > 0) {
      aiResponse = json.candidates[0].content.parts[0].text;
    } else {
      aiResponse = "Gemini APIエラー: 応答の解析に失敗しました。詳細: " + JSON.stringify(json);
    }

    // 質問と回答のログをDBに蓄積（将来の自己学習や組織の疑問点分析のため）
    appendToSheet('qalogs', [getCurrentTime(), query, aiResponse]);

    return aiResponse;
  } catch (e) {
    return "通信エラーが発生しました。ネットワーク設定またはAPI制限を確認してください: " + e.message;
  }
}
