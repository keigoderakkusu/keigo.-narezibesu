// ============================================================
// Feature_Approval.gs  ── ⑫ 承認ワークフロー
// ナレッジ記事を「申請 → 承認 / 差し戻し → 公開」フローで管理
// 承認リクエストはGmailで通知、スプレッドシートで追跡
// ============================================================

const PROP_APPROVER_EMAIL = 'APPROVER_EMAIL';   // 承認者のGmailアドレス
const PROP_APP_URL        = 'APP_URL';           // WebアプリのURL（通知リンク用）

// ── シート拡張 ──────────────────────────────────────────────
function setupApprovalDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = {
    'approval_requests': [
      'リクエストID','記事ID','記事タイトル','申請者','申請日時',
      'ステータス(pending/approved/rejected)','承認者','承認日時','コメント'
    ],
    'approval_logs': [
      'ログID','リクエストID','操作','操作者','日時','備考'
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
  return '承認DBの初期化完了';
}

// ── 承認リクエスト送信 ──────────────────────────────────────
/**
 * KB記事の承認を申請する
 * ステータスを「承認待ち」に変更し、承認者にGmailで通知
 */
function requestApproval(articleId, requesterName) {
  const arts    = rows('kb_articles');
  const article = arts.find(a => String(a['記事ID']) === String(articleId));
  if (!article) return { status: 'error', message: '記事が見つかりません' };

  const reqId     = uid();
  const appUrl    = PropertiesService.getScriptProperties().getProperty(PROP_APP_URL) || '';
  const approverEmail = PropertiesService.getScriptProperties().getProperty(PROP_APPROVER_EMAIL);

  // ステータスを「承認待ち」へ更新
  updateRowData('kb_articles', articleId, 'ステータス(公開/下書き)', '承認待ち');

  // 承認リクエスト登録
  push('approval_requests', [
    reqId, articleId, article['タイトル'],
    requesterName || Session.getActiveUser().getEmail(),
    now(), 'pending', '', '', ''
  ]);

  // 承認ログ
  push('approval_logs', [uid(), reqId, '申請', requesterName||'作成者', now(), article['タイトル']]);

  // Gmail通知
  if (approverEmail) {
    try {
      const subject = `【承認リクエスト】${article['タイトル']}`;
      const body = buildApprovalEmailBody(article, reqId, appUrl, requesterName, 'request');
      GmailApp.sendEmail(approverEmail, subject, '', { htmlBody: body });
    } catch(e) { Logger.log('メール送信失敗: ' + e.message); }
  }

  return {
    status: 'success',
    requestId: reqId,
    message: `「${article['タイトル']}」の承認リクエストを送信しました。`
  };
}

// ── 承認処理 ────────────────────────────────────────────────
/**
 * 承認リクエストを承認する
 * KB記事を「公開」に変更し、申請者にGmailで通知
 */
function approveRequest(requestId, approverName, comment) {
  const reqs = rows('approval_requests');
  const req  = reqs.find(r => String(r['リクエストID']) === String(requestId));
  if (!req) return { status: 'error', message: 'リクエストが見つかりません' };
  if (req['ステータス(pending/approved/rejected)'] !== 'pending') {
    return { status: 'error', message: 'すでに処理済みのリクエストです' };
  }

  const articleId = req['記事ID'];

  // リクエストを更新
  updateRowData('approval_requests', requestId, 'ステータス(pending/approved/rejected)', 'approved');
  updateRowData('approval_requests', requestId, '承認者', approverName || Session.getActiveUser().getEmail());
  updateRowData('approval_requests', requestId, '承認日時', now());
  updateRowData('approval_requests', requestId, 'コメント', comment || '');

  // KB記事を公開
  updateRowData('kb_articles', articleId, 'ステータス(公開/下書き)', '公開');

  // ログ
  push('approval_logs', [uid(), requestId, '承認', approverName||'承認者', now(), comment||'']);

  // 申請者へ通知
  try {
    const requesterEmail = req['申請者'];
    if (requesterEmail && requesterEmail.includes('@')) {
      const subject = `【承認完了】${req['記事タイトル']}`;
      const body = buildApprovalEmailBody({ タイトル: req['記事タイトル'], 記事ID: articleId },
        requestId, PropertiesService.getScriptProperties().getProperty(PROP_APP_URL)||'', approverName, 'approved', comment);
      GmailApp.sendEmail(requesterEmail, subject, '', { htmlBody: body });
    }
  } catch(e) { Logger.log('承認通知メール失敗: ' + e.message); }

  return { status: 'success', message: `「${req['記事タイトル']}」を承認し公開しました。` };
}

// ── 差し戻し処理 ────────────────────────────────────────────
/**
 * 承認リクエストを差し戻す
 * KB記事を「下書き」に戻し、申請者にGmailで理由を通知
 */
function rejectRequest(requestId, approverName, reason) {
  const reqs = rows('approval_requests');
  const req  = reqs.find(r => String(r['リクエストID']) === String(requestId));
  if (!req) return { status: 'error', message: 'リクエストが見つかりません' };

  updateRowData('approval_requests', requestId, 'ステータス(pending/approved/rejected)', 'rejected');
  updateRowData('approval_requests', requestId, '承認者', approverName || '');
  updateRowData('approval_requests', requestId, '承認日時', now());
  updateRowData('approval_requests', requestId, 'コメント', reason || '');

  // KB記事を下書きに戻す
  updateRowData('kb_articles', req['記事ID'], 'ステータス(公開/下書き)', '下書き');

  push('approval_logs', [uid(), requestId, '差し戻し', approverName||'承認者', now(), reason||'']);

  // 申請者へ通知
  try {
    const requesterEmail = req['申請者'];
    if (requesterEmail && requesterEmail.includes('@')) {
      GmailApp.sendEmail(requesterEmail, `【差し戻し】${req['記事タイトル']}`, '',
        { htmlBody: buildApprovalEmailBody({ タイトル: req['記事タイトル'] }, requestId, '', approverName, 'rejected', reason) });
    }
  } catch(e) {}

  return { status: 'success', message: `「${req['記事タイトル']}」を差し戻しました。` };
}

// ── 承認待ちリスト取得 ────────────────────────────────────
function getPendingApprovals() {
  return rows('approval_requests')
    .filter(r => r['ステータス(pending/approved/rejected)'] === 'pending')
    .sort((a, b) => new Date(b['申請日時']) - new Date(a['申請日時']));
}

// ── 承認履歴取得 ─────────────────────────────────────────
function getApprovalHistory(limit) {
  const all = rows('approval_requests')
    .sort((a, b) => new Date(b['申請日時']) - new Date(a['申請日時']));
  return limit ? all.slice(0, limit) : all;
}

// ── 自分の申請リスト取得 ────────────────────────────────────
function getMyRequests(requester) {
  return rows('approval_requests')
    .filter(r => r['申請者'] === requester)
    .sort((a, b) => new Date(b['申請日時']) - new Date(a['申請日時']));
}

// ── メールテンプレート ───────────────────────────────────────
function buildApprovalEmailBody(article, requestId, appUrl, actor, type, comment) {
  const colors = { request:'#2563eb', approved:'#16a34a', rejected:'#dc2626' };
  const labels = { request:'承認リクエスト', approved:'承認完了 ✅', rejected:'差し戻し ❌' };
  const color  = colors[type] || '#64748b';
  const label  = labels[type] || type;

  return `
  <!DOCTYPE html>
  <html>
  <head><meta charset="utf-8"></head>
  <body style="font-family:sans-serif;background:#f8fafc;padding:24px;color:#0f172a;">
    <div style="max-width:560px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08);">
      <div style="background:${color};padding:20px 24px;">
        <h2 style="color:#fff;margin:0;font-size:1.2rem;">【${label}】Sales Knowledge Hub</h2>
      </div>
      <div style="padding:24px;">
        <p style="margin-bottom:16px;font-size:14px;line-height:1.7;">
          ${type==='request'
            ? `<strong>${actor || '申請者'}</strong> さんから記事の承認リクエストが届きました。`
            : type==='approved'
            ? `<strong>${actor || '承認者'}</strong> さんが記事を承認しました。`
            : `<strong>${actor || '承認者'}</strong> さんが記事を差し戻しました。`}
        </p>
        <div style="background:#f1f5f9;border-radius:8px;padding:16px;margin-bottom:20px;">
          <div style="font-size:11px;color:#64748b;font-weight:700;text-transform:uppercase;margin-bottom:6px;">記事タイトル</div>
          <div style="font-size:16px;font-weight:700;">${article['タイトル'] || ''}</div>
        </div>
        ${comment ? `<div style="border-left:3px solid ${color};padding:12px 16px;margin-bottom:20px;background:#f8fafc;font-size:14px;"><strong>コメント:</strong> ${comment}</div>` : ''}
        ${appUrl && type==='request' ? `<a href="${appUrl}" style="display:inline-block;background:${color};color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px;">アプリを開いて確認する →</a>` : ''}
        <p style="margin-top:24px;font-size:12px;color:#94a3b8;">このメールは Sales Knowledge Hub から自動送信されました。</p>
      </div>
    </div>
  </body>
  </html>`;
}

// ── 定期チェック（承認忘れリマインダー） ──────────────────────
// GASのトリガーで毎日実行
function sendApprovalReminders() {
  const pending = getPendingApprovals();
  if (!pending.length) return '承認待ちリクエストなし';

  const approverEmail = PropertiesService.getScriptProperties().getProperty(PROP_APPROVER_EMAIL);
  if (!approverEmail) return '承認者メール未設定';

  // 24時間以上経過したリクエストをリマインド
  const old = pending.filter(r => {
    const h = (new Date() - new Date(r['申請日時'])) / 3600000;
    return h > 24;
  });
  if (!old.length) return 'リマインド不要';

  const list = old.map(r => `<li style="margin-bottom:8px;"><strong>${r['記事タイトル']}</strong> (申請: ${r['申請者']}, ${r['申請日時']})</li>`).join('');
  const body = `
    <html><body style="font-family:sans-serif;padding:24px;">
      <h2 style="color:#d97706;">⏰ 承認待ちリマインダー</h2>
      <p>以下の記事が24時間以上承認待ちです:</p>
      <ul>${list}</ul>
      <a href="${PropertiesService.getScriptProperties().getProperty(PROP_APP_URL)||''}" style="display:inline-block;background:#2563eb;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-weight:600;">アプリを開く</a>
    </body></html>`;

  GmailApp.sendEmail(approverEmail, `【リマインド】${old.length}件の承認待ちリクエストがあります`, '', { htmlBody: body });
  return `${old.length}件のリマインドメールを送信しました`;
}
