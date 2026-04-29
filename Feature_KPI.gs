// ============================================================
// Feature_KPI.gs  ── ⑨ KPI進捗ダッシュボード
// 月次目標に対する達成率をリアルタイムで計算
// Googleカレンダー予定と連動して「今月残り予定商談数」も表示
// ============================================================

// ── シート拡張 ──────────────────────────────────────────────
function setupKpiDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = {
    'kpi_targets': [
      'KPI_ID','年月(YYYY/MM)','KPI名','目標値','単位',
      'カテゴリ(商談/ナレッジ/タスク/売上/その他)','備考'
    ],
    'kpi_snapshots': [
      'スナップショットID','取得日時','年月','KPI名','実績値','達成率(%)'
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
  return 'KPI DBの初期化完了';
}

// ── KPI目標の登録 ────────────────────────────────────────────
function saveKpiTarget(d) {
  const ym = d.yearMonth || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM');
  push('kpi_targets', [uid(), ym, d.name, parseFloat(d.target)||0, d.unit||'件', d.category||'その他', d.note||'']);
  return { status: 'success', message: `KPI目標「${d.name}」を登録しました。` };
}

function updateKpiTarget(id, d) {
  ['KPI名','目標値','単位','カテゴリ(商談/ナレッジ/タスク/売上/その他)','備考'].forEach(k => {
    if (d[k] !== undefined) updateRowData('kpi_targets', id, k, d[k]);
  });
  return '更新しました';
}

function deleteKpiTarget(id) { return deleteRowData('kpi_targets', id); }

// ── メイン: 今月のKPIダッシュボードデータ取得 ─────────────────
/**
 * 今月の目標 vs 実績を計算して返す
 * カレンダーの予定商談数も含める
 */
function getKpiDashboard(yearMonth) {
  const ym = yearMonth || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM');
  const [year, month] = ym.split('/').map(Number);

  const monthStart = new Date(year, month-1, 1);
  const monthEnd   = new Date(year, month, 0, 23, 59, 59);
  const today      = new Date();

  // ── 実績計算 ──────────────────────────────────────────────

  // 商談件数（今月登録）
  const meetings = rows('meetings').filter(m => {
    const d = new Date(m['登録日時']);
    return d >= monthStart && d <= monthEnd;
  });

  // ナレッジ作成数（KB記事 + メモ）
  const kbArts = rows('kb_articles').filter(a => {
    const d = new Date(a['登録日時']);
    return d >= monthStart && d <= monthEnd && a['ステータス(公開/下書き)'] === '公開';
  });
  const notes = rows('notes').filter(n => {
    const d = new Date(n['登録日時']);
    return d >= monthStart && d <= monthEnd;
  });

  // タスク稼働時間（今月）
  const tasks = rows('tasks').filter(t => {
    const d = new Date(t['登録日時']);
    return d >= monthStart && d <= monthEnd && parseInt(t['所要時間(分)']) > 0;
  });
  const taskTotalMin = tasks.reduce((s,t) => s + (parseInt(t['所要時間(分)'])||0), 0);

  // マニュアル作成数
  const manuals = rows('manuals').filter(m => {
    const d = new Date(m['登録日時']);
    return d >= monthStart && d <= monthEnd;
  });

  // ── カレンダー予定取得 ────────────────────────────────────
  let calendarMeetings = [], upcomingMeetings = [];
  try {
    const events = CalendarApp.getDefaultCalendar().getEvents(monthStart, monthEnd);
    const meetingKeywords = ['訪問','商談','打合','ミーティング','meeting','client','顧客','提案'];
    calendarMeetings = events.filter(e =>
      meetingKeywords.some(k => e.getTitle().toLowerCase().includes(k.toLowerCase()))
    ).map(e => ({
      title: e.getTitle(),
      start: Utilities.formatDate(e.getStartTime(), Session.getScriptTimeZone(), 'MM/dd HH:mm'),
      isPast: e.getStartTime() < today
    }));
    upcomingMeetings = calendarMeetings.filter(e => !e.isPast);
  } catch(e) { Logger.log('カレンダー取得失敗: ' + e.message); }

  // ── KPI目標取得 ──────────────────────────────────────────
  const targets = rows('kpi_targets').filter(t => t['年月(YYYY/MM)'] === ym || t['年月(YYYY/MM)'] === '毎月');

  // ── 実績マップ ────────────────────────────────────────────
  const actuals = {
    '商談件数':           meetings.length,
    'ナレッジ作成数':     kbArts.length + notes.length,
    'KB記事公開数':       kbArts.length,
    'タスク稼働時間(h)':  Math.round(taskTotalMin / 60 * 10) / 10,
    'マニュアル作成数':   manuals.length,
    'カレンダー商談予定': calendarMeetings.length,
  };

  // ── KPIカード生成 ─────────────────────────────────────────
  const kpiCards = targets.map(t => {
    const name   = t['KPI名'];
    const target = parseFloat(t['目標値']) || 0;
    const actual = actuals[name] !== undefined ? actuals[name] : 0;
    const rate   = target > 0 ? Math.min(Math.round(actual / target * 100), 150) : 0;

    let status = 'normal';
    if (rate >= 100) status = 'achieved';
    else if (rate >= 70) status = 'on_track';
    else if (rate < 40 && today.getDate() > 15) status = 'behind';

    return { id: t['KPI_ID'], name, target, actual, unit: t['単位']||'件', rate, status, category: t['カテゴリ(商談/ナレッジ/タスク/売上/その他)'] };
  });

  // ── デフォルトKPIカード（目標未設定でも表示） ─────────────
  const defaultKpis = [
    { name:'商談件数', unit:'件', category:'商談' },
    { name:'ナレッジ作成数', unit:'件', category:'ナレッジ' },
    { name:'タスク稼働時間(h)', unit:'h', category:'タスク' },
    { name:'マニュアル作成数', unit:'件', category:'ナレッジ' },
  ];
  defaultKpis.forEach(dk => {
    if (!kpiCards.find(k => k.name === dk.name)) {
      kpiCards.push({
        id: null, name: dk.name, target: 0, actual: actuals[dk.name] || 0,
        unit: dk.unit, rate: 0, status: 'no_target', category: dk.category
      });
    }
  });

  // ── 月次トレンド（過去6ヶ月） ─────────────────────────────
  const trend = buildMonthlyTrend(6);

  // ── 今月の日別ペース ──────────────────────────────────────
  const dailyPace = buildDailyPace(meetings, year, month);

  // ── スナップショット保存（定期記録用） ─────────────────────
  kpiCards.forEach(k => {
    if (k.actual > 0) {
      push('kpi_snapshots', [uid(), now(), ym, k.name, k.actual, k.rate]);
    }
  });

  return {
    yearMonth:        ym,
    kpiCards,
    actuals,
    calendarMeetings,
    upcomingCount:    upcomingMeetings.length,
    completedCount:   calendarMeetings.length - upcomingMeetings.length,
    trend,
    dailyPace,
    summary: {
      totalMeetings:   meetings.length,
      totalKnowledge:  kbArts.length + notes.length,
      totalTaskHours:  Math.round(taskTotalMin / 60 * 10) / 10,
      upcomingMeetings
    }
  };
}

// ── 過去6ヶ月のトレンドデータ ─────────────────────────────────
function buildMonthlyTrend(months) {
  const trend = { labels:[], meetings:[], knowledge:[], taskHours:[] };
  const now   = new Date();

  for (let i = months - 1; i >= 0; i--) {
    const d     = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const year  = d.getFullYear();
    const month = d.getMonth() + 1;
    const ym    = `${year}/${String(month).padStart(2,'0')}`;
    const start = new Date(year, month-1, 1);
    const end   = new Date(year, month, 0, 23, 59, 59);

    trend.labels.push(ym.replace(/\d{4}\//,''));

    const mtgCount = rows('meetings').filter(m => {
      const d2 = new Date(m['登録日時']);
      return d2 >= start && d2 <= end;
    }).length;

    const knCount = rows('kb_articles').filter(a => {
      const d2 = new Date(a['登録日時']);
      return d2 >= start && d2 <= end;
    }).length + rows('notes').filter(n => {
      const d2 = new Date(n['登録日時']);
      return d2 >= start && d2 <= end;
    }).length;

    const taskMin = rows('tasks').filter(t => {
      const d2 = new Date(t['登録日時']);
      return d2 >= start && d2 <= end;
    }).reduce((s,t) => s + (parseInt(t['所要時間(分)'])||0), 0);

    trend.meetings.push(mtgCount);
    trend.knowledge.push(knCount);
    trend.taskHours.push(Math.round(taskMin/60*10)/10);
  }
  return trend;
}

// ── 今月の日別商談ペース ─────────────────────────────────────
function buildDailyPace(meetings, year, month) {
  const daysInMonth = new Date(year, month, 0).getDate();
  const daily = Array(daysInMonth).fill(0);

  meetings.forEach(m => {
    const d = new Date(m['登録日時']);
    if (d.getMonth() + 1 === month) {
      daily[d.getDate() - 1]++;
    }
  });

  const labels = Array.from({length: daysInMonth}, (_, i) => `${i+1}日`);
  return { labels, values: daily };
}

// ── KPI目標一覧取得 ──────────────────────────────────────────
function getKpiTargets(yearMonth) {
  const ym = yearMonth || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM');
  return rows('kpi_targets').filter(t => t['年月(YYYY/MM)'] === ym || t['年月(YYYY/MM)'] === '毎月');
}

// ── 月次サマリレポートをメール/Chat送信 ──────────────────────
// GASトリガーで月初に実行
function sendMonthlyKpiReport() {
  const lastMonth = (() => {
    const d = new Date(); d.setMonth(d.getMonth() - 1);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM');
  })();

  const data     = getKpiDashboard(lastMonth);
  const achieved = data.kpiCards.filter(k => k.status === 'achieved').length;
  const total    = data.kpiCards.filter(k => k.status !== 'no_target').length;

  const rows_ = data.kpiCards
    .filter(k => k.status !== 'no_target')
    .map(k => `<tr><td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;">${k.name}</td><td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;text-align:right;">${k.target}${k.unit}</td><td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;text-align:right;font-weight:700;color:${k.rate>=100?'#16a34a':k.rate>=70?'#d97706':'#dc2626'};">${k.actual}${k.unit}</td><td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;text-align:right;">${k.rate}%</td></tr>`)
    .join('');

  const html = `
  <!DOCTYPE html><html><head><meta charset="utf-8"></head>
  <body style="font-family:sans-serif;background:#f8fafc;padding:24px;color:#0f172a;">
    <div style="max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.08);">
      <div style="background:linear-gradient(135deg,#1e293b,#0f172a);padding:24px;">
        <h1 style="color:#fff;margin:0;font-size:1.3rem;">📊 月次KPIレポート ${lastMonth}</h1>
        <p style="color:#94a3b8;margin:8px 0 0;font-size:13px;">Sales Knowledge Hub 自動レポート</p>
      </div>
      <div style="padding:24px;">
        <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:16px;margin-bottom:20px;text-align:center;">
          <div style="font-size:2rem;font-weight:700;color:#16a34a;">${achieved} / ${total}</div>
          <div style="font-size:13px;color:#166534;">KPI達成</div>
        </div>
        <table style="width:100%;border-collapse:collapse;font-size:13.5px;">
          <thead><tr style="background:#f1f5f9;">
            <th style="padding:8px 12px;text-align:left;font-weight:700;color:#64748b;">KPI</th>
            <th style="padding:8px 12px;text-align:right;font-weight:700;color:#64748b;">目標</th>
            <th style="padding:8px 12px;text-align:right;font-weight:700;color:#64748b;">実績</th>
            <th style="padding:8px 12px;text-align:right;font-weight:700;color:#64748b;">達成率</th>
          </tr></thead>
          <tbody>${rows_}</tbody>
        </table>
        <div style="margin-top:20px;padding:16px;background:#f8fafc;border-radius:8px;font-size:13px;color:#475569;">
          <strong>今月の予定商談:</strong> ${data.upcomingCount}件<br>
          <strong>タスク稼働時間:</strong> ${data.actuals['タスク稼働時間(h)']}h
        </div>
        <p style="margin-top:20px;font-size:12px;color:#94a3b8;">このレポートは Sales Knowledge Hub から自動送信されました。</p>
      </div>
    </div>
  </body></html>`;

  const approverEmail = PropertiesService.getScriptProperties().getProperty(PROP_APPROVER_EMAIL);
  const myEmail = Session.getActiveUser().getEmail();
  const to = myEmail || approverEmail;
  if (to) {
    GmailApp.sendEmail(to, `📊 月次KPIレポート ${lastMonth}`, '', { htmlBody: html });
  }

  const webhook = PropertiesService.getScriptProperties().getProperty(PROP_WEBHOOK);
  if (webhook) {
    const text = `📊 *${lastMonth} 月次KPIレポート*\n達成: ${achieved}/${total} KPI\n商談: ${data.actuals['商談件数']}件 / ナレッジ: ${data.actuals['ナレッジ作成数']}件 / タスク: ${data.actuals['タスク稼働時間(h)']}h`;
    UrlFetchApp.fetch(webhook, { method:'post', contentType:'application/json', payload: JSON.stringify({text}), muteHttpExceptions:true });
  }

  return `月次KPIレポートを送信しました（${lastMonth}）`;
}
