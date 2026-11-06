/**
 * 英語勉強会ボード メール送信係(Google Apps Script)
 *
 * 1分ごとに Firestore を確認して、次のメールを送ります。
 *  ・新しい予定が追加されたとき → メンバー全員(メール通知オンの人)にお知らせ
 *  ・誰かが「参加」を押したとき → その人に予約確定メール(カレンダー用 .ics 付き)
 *  ・毎晩20時 → 明日の勉強会のリマインド(参加/未定/未返事で文面が変わる。不参加の人には送らない)
 *  ・日曜19時 → 週まとめ(来週の予定・先週のアウトプット・ストリーク・文法の進み具合・人気テーマ)
 *
 * 初回だけ: 関数「setup」を選んで ▶実行 → Googleの許可画面で許可。
 * お試し: 「testReminder」「testDigest」を実行すると、このアカウント自身にだけ見本が届きます。
 * 送信元は、このスクリプトを動かしている Google アカウントの Gmail です。
 */
const PROJECT = 'study-board-9aefe';
const APP_URL = 'https://shintaroyade.github.io/study-board/';
const DEFAULT_NAME = '英語勉強会';
const DEFAULT_INVITE = 'OUTPUT2026';
const ADMIN_UID = 'L8IGhzLvCqVWjKxoLXQgDYsc8lH3';
const SENDER = 'output.apu2026@gmail.com';   // このスクリプトを動かしているアカウント(お試し送信の宛先)
const REMIND_HOUR = 20;               // 前日リマインドを送る時刻(時)
const DIGEST_DOW = 7, DIGEST_HOUR = 19; // 週まとめ:日曜(7) の 19時
const BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents/';
const TYPES = { grammar: '英文法', lesson: 'メンバー授業', other: 'その他' };
const UNITS = { bunkei:'文型', hinshi:'品詞', jisei:'時制', kanryo:'完了形', jodoushi:'助動詞', ukemi:'受動態', futeishi:'不定詞', doumeishi:'動名詞', bunshi:'分詞', hikaku:'比較', kankei:'関係詞', katei:'仮定法', wahou:'話法', meishi:'名詞・冠詞', zenchi:'前置詞', setsuzoku:'接続詞', tokushu:'特殊構文' };
const UNIT_COUNT = Object.keys(UNITS).length;
const WD = ['日','月','火','水','木','金','土'];

/** 初回に1回だけ実行してください */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('since')) p.setProperty('since', String(Date.now()));
  // 招待コードがまだなければ初期値を入れる
  if (!fsGet('config/invite')) fsSet('config/invite', { code: DEFAULT_INVITE });
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(1).create();
  const n = fsList('events').length;
  Logger.log('準備OK:予定 ' + n + ' 件を確認できました。これから1分ごとにチェックしてメールを送ります。');
}

/** データをまとめて読み込む */
function loadCtx() {
  const settings = fsGet('settings/app');
  const c = { appName: (settings && settings.name) || DEFAULT_NAME, members: {}, contacts: {}, events: {}, rsvps: [] };
  fsList('members').forEach(function (m) { c.members[m.id] = m; });
  fsList('contacts').forEach(function (x) { c.contacts[x.id] = x; });
  fsList('events').forEach(function (e) { c.events[e.id] = e; });
  c.rsvps = fsList('rsvps');
  c.wants = function (uid) { const x = c.contacts[uid]; return x && x.email && x.notify !== false && c.members[uid]; };
  c.rsvpOf = function (eid, uid) { for (var i = 0; i < c.rsvps.length; i++) { const r = c.rsvps[i]; if (r.eventId === eid && r.uid === uid) return r.status; } return ''; };
  return c;
}

/** 1分ごとに自動で動く */
function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const since = Number(props.getProperty('since') || Date.now());
    const c = loadCtx();

    // 1) 新しい予定 → 全員にお知らせ
    Object.keys(c.events).forEach(function (id) {
      const e = c.events[id]; const key = 'ev_' + id;
      if (props.getProperty(key) || !(Number(e.createdAt) >= since)) return;
      props.setProperty(key, String(Date.now()));
      const by = c.members[e.createdBy] ? c.members[e.createdBy].name : 'メンバー';
      Object.keys(c.contacts).forEach(function (uid) {
        if (!c.wants(uid)) return;
        MailApp.sendEmail({
          to: c.contacts[uid].email,
          subject: '【' + c.appName + '】新しい予定:' + md(e.date) + ' ' + title(e),
          name: c.appName,
          htmlBody: wrap(c.appName, (c.members[uid].name || '') + 'さん、' + by + 'さんが新しい予定を追加しました。',
            eventBlock(e, c.members, true) + '<a href="' + APP_URL + '" style="' + BTN + '">アプリで参加を返事する</a>'),
        });
      });
    });

    // 2) 「参加」→ 予約確定メール
    c.rsvps.forEach(function (r) {
      if (r.status !== 'yes') return;
      const key = 'rs_' + r.id;
      if (props.getProperty(key) || !(Number(r.at) >= since)) return;
      props.setProperty(key, String(Date.now()));
      const e = c.events[r.eventId];
      if (!e || !c.wants(r.uid)) return;
      MailApp.sendEmail({
        to: c.contacts[r.uid].email,
        subject: '【予約確定】' + md(e.date) + ' ' + title(e) + '|' + c.appName,
        name: c.appName,
        htmlBody: wrap(c.appName, (c.members[r.uid].name || '') + 'さん、参加の予約が確定しました。',
          eventBlock(e, c.members, true) + '<a href="' + gcalUrl(e, c.appName) + '" style="' + BTN + '">Googleカレンダーに追加</a>' +
          '<p style="font-size:12px;color:#5B6782;margin:10px 0 0">iPhoneなどは、添付の「study.ics」を開くとカレンダーに追加できます。</p>'),
        attachments: [Utilities.newBlob(ics(e, c.appName), 'text/calendar', 'study.ics')],
      });
    });

    // 3) 毎晩20時:明日のリマインド
    const n = nowJst();
    if (n.hour >= REMIND_HOUR && !props.getProperty('remind_' + n.date)) {
      props.setProperty('remind_' + n.date, String(Date.now()));
      sendReminders(addDays(n.date, 1), c);
    }
    // 4) 日曜19時:週まとめ
    if (n.dow === DIGEST_DOW && n.hour >= DIGEST_HOUR && !props.getProperty('digest_' + n.date)) {
      props.setProperty('digest_' + n.date, String(Date.now()));
      sendDigest(n.date, c);
    }

    // 古い送信記録を掃除(60日)
    const cut = Date.now() - 60 * 864e5; const all = props.getProperties();
    Object.keys(all).forEach(function (k) { if (/^(ev|rs|remind|digest)_/.test(k) && Number(all[k]) < cut) props.deleteProperty(k); });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- 前日リマインド ---------------- */
function sendReminders(day, c, testTo) {
  const evs = Object.keys(c.events).map(function (k) { return c.events[k]; })
    .filter(function (e) { return e.date === day; }).sort(byTime);
  if (!evs.length) return 0;
  let sent = 0;
  const uids = testTo ? [pickTestUid(c)] : Object.keys(c.contacts);
  uids.forEach(function (uid) {
    if (!uid || (!testTo && !c.wants(uid))) return;
    const blocks = [];
    let needReply = false;
    evs.forEach(function (e) {
      const st = c.rsvpOf(e.id, uid);
      if (st === 'no') return;
      let note;
      if (st === 'yes') note = '<p style="' + NOTE + 'background:#EEFAD0;color:#365300">✅ 参加で返事済みです。待ってます!</p>';
      else if (st === 'maybe') { note = '<p style="' + NOTE + 'background:#FFF4D6;color:#6B4E00">🤔 「未定」のままです。行けるか決まったらアプリで返事してね。</p>'; needReply = true; }
      else { note = '<p style="' + NOTE + 'background:#E6F7FB;color:#0B5566">📩 まだ返事がありません。参加/不参加をアプリで押してね。</p>'; needReply = true; }
      const prep = e.memo ? '<p style="' + NOTE + 'background:#F2F4F8;color:#0F172A"><b>📝 予習・準備</b><br>' + h(e.memo).replace(/\n/g, '<br>') + '</p>' : '';
      blocks.push(eventBlock(e, c.members) + prep + note + (e.link ? '<p style="font-size:13px;margin:0 0 14px">📎 資料:<a href="' + h(e.link) + '">' + h(e.link) + '</a></p>' : ''));
    });
    if (!blocks.length) return;
    const name = c.members[uid] ? c.members[uid].name : '';
    MailApp.sendEmail({
      to: testTo || c.contacts[uid].email,
      subject: (testTo ? '[テスト] ' : '') + '【明日】' + md(day) + ' ' + evs.map(title).join('・') + '|' + c.appName,
      name: c.appName,
      htmlBody: wrap(c.appName, name + 'さん、明日は勉強会があります。', blocks.join('<hr style="border:0;border-top:1px solid #E4EAF4;margin:18px 0">') +
        '<a href="' + APP_URL + '" style="' + BTN + '">' + (needReply ? 'アプリで返事する' : 'アプリを開く') + '</a>'),
    });
    sent++;
  });
  return sent;
}

/* ---------------- 日曜の週まとめ ---------------- */
function sendDigest(today, c, testTo) {
  const from = addDays(today, 1), to = addDays(today, 7);
  const next = Object.keys(c.events).map(function (k) { return c.events[k]; })
    .filter(function (e) { return e.date >= from && e.date <= to; }).sort(byTime);
  const weekAgo = Date.now() - 7 * 864e5;
  const outs = fsList('outputs').filter(function (o) { return o.learned && Number(o.at) >= weekAgo; })
    .sort(function (a, b) { return Number(b.at) - Number(a.at); });
  const st = stats(c, today);
  const lastWeekAtt = Object.keys(c.events).map(function (k) { return c.events[k]; })
    .filter(function (e) { return e.date > addDays(today, -7) && e.date <= today; })
    .reduce(function (s, e) { return s + ((e.attended || []).length); }, 0);
  const prog = fsGet('grammar/progress');
  const done = prog && prog.statuses ? Object.keys(prog.statuses).filter(function (k) { return prog.statuses[k] === 'done'; }).length : 0;
  const themes = fsList('themes').filter(function (t) { return t.status !== 'scheduled'; })
    .sort(function (a, b) { return (b.votes || []).length - (a.votes || []).length; }).slice(0, 3);
  if (!next.length && !outs.length && !lastWeekAtt) return 0;

  const top = Object.keys(st).filter(function (id) { return st[id].streak > 0; })
    .sort(function (a, b) { return st[b].streak - st[a].streak || st[b].count - st[a].count; }).slice(0, 3);
  const sec = function (t) { return '<p style="font-size:15px;font-weight:700;margin:22px 0 8px;padding-top:14px;border-top:1px solid #E4EAF4">' + t + '</p>'; };

  let sent = 0;
  const uids = testTo ? [pickTestUid(c)] : Object.keys(c.contacts);
  uids.forEach(function (uid) {
    if (!uid || (!testTo && !c.wants(uid))) return;
    const me = st[uid] || { streak: 0, count: 0 };
    const name = c.members[uid] ? c.members[uid].name : '';
    let body = '<table style="width:100%;border-collapse:separate;border-spacing:6px 0;margin:0 -6px 6px"><tr>' +
      stat('🔥' + me.streak, 'あなたの連続参加') + stat(String(me.count), 'あなたの通算') + stat(lastWeekAtt + '人', '先週のみんなの出席') + '</tr></table>';
    if (me.streak > 0) body += '<p style="font-size:13px;color:#365300;margin:6px 0 0">次に参加すれば ' + (me.streak + 1) + ' 連続!</p>';
    else body += '<p style="font-size:13px;color:#5B6782;margin:6px 0 0">来週の1回でストリークがスタートします。</p>';

    body += sec('📅 来週の予定(' + md(from) + '〜' + md(to) + ')');
    if (next.length) {
      body += next.map(function (e) {
        const s = c.rsvpOf(e.id, uid);
        const badge = s === 'yes' ? '<span style="color:#365300">✅ 参加</span>' : s === 'no' ? '<span style="color:#8C9AB8">不参加</span>' : s === 'maybe' ? '<span style="color:#6B4E00">🤔 未定</span>' : '<span style="color:#0B5566">📩 未返事</span>';
        const host = c.members[e.hostId] ? ' / 担当 ' + h(c.members[e.hostId].name) : '';
        return '<div style="padding:10px 12px;border:1px solid #E4EAF4;border-radius:10px;margin-bottom:8px;font-size:14px">' +
          '<b>' + md(e.date) + ' ' + h(timeRange(e)) + '</b> ' + badge + '<br>' + h(TYPES[e.type] || '') + ':' + h(title(e)) + host +
          (e.place ? '<br><span style="color:#5B6782">📍 ' + h(e.place) + '</span>' : '') + '</div>';
      }).join('');
    } else body += '<p style="font-size:14px;color:#5B6782;margin:0">来週の予定はまだありません。アプリから追加しよう。</p>';

    if (outs.length) {
      body += sec('✍️ 先週のみんなのアウトプット');
      body += outs.slice(0, 8).map(function (o) {
        const e = c.events[o.eventId]; const who = c.members[o.uid] ? c.members[o.uid].name : 'メンバー';
        return '<div style="font-size:14px;margin-bottom:10px"><b>' + h(who) + '</b>' + (e ? '<span style="color:#5B6782;font-size:12px"> ・' + md(e.date) + ' ' + h(title(e)) + '</span>' : '') +
          '<br>「' + h(o.learned).replace(/\n/g, '<br>') + '」</div>';
      }).join('');
    }

    if (top.length) {
      body += sec('🔥 連続参加ランキング');
      body += top.map(function (id, i) { return '<div style="font-size:14px;margin-bottom:4px">' + ['🥇', '🥈', '🥉'][i] + ' ' + h(c.members[id].name) + ' さん ― ' + st[id].streak + ' 連続</div>'; }).join('');
    }

    body += sec('📖 英文法の進み具合');
    const pct = Math.round(done / UNIT_COUNT * 100);
    body += '<p style="font-size:14px;margin:0 0 6px">' + done + ' / ' + UNIT_COUNT + ' 単元 完了(' + pct + '%)</p>' +
      '<div style="height:10px;border-radius:99px;background:#E4EAF4"><div style="height:10px;border-radius:99px;background:#22D3EE;width:' + pct + '%"></div></div>';

    if (themes.length) {
      body += sec('💡 人気の「聞きたいテーマ」');
      body += themes.map(function (t) { return '<div style="font-size:14px;margin-bottom:4px">👍' + (t.votes || []).length + ' ' + h(t.title) + (t.teacher && c.members[t.teacher] ? ' <span style="color:#C2410C">(' + h(c.members[t.teacher].name) + 'さんが教えられる)</span>' : '') + '</div>'; }).join('');
      body += '<p style="font-size:12px;color:#5B6782;margin:6px 0 0">教えられそうなテーマがあれば、アプリで「🙋 教えられる」を押してね。</p>';
    }
    body += '<p style="margin:20px 0 0"><a href="' + APP_URL + '" style="' + BTN + '">アプリを開く</a></p>';

    MailApp.sendEmail({
      to: testTo || c.contacts[uid].email,
      subject: (testTo ? '[テスト] ' : '') + '【今週のまとめ】来週の予定 ' + next.length + '件|' + c.appName,
      name: c.appName,
      htmlBody: wrap(c.appName, name + 'さん、今週もおつかれさま!1週間のまとめです。', body),
    });
    sent++;
  });
  return sent;
}

/* ---------------- お試し送信(このアカウント自身にだけ届く) ---------------- */
function testReminder() {
  const c = loadCtx(); const me = SENDER; const today = nowJst().date;
  let day = addDays(today, 1);
  if (!Object.keys(c.events).some(function (k) { return c.events[k].date === day; })) {
    const up = Object.keys(c.events).map(function (k) { return c.events[k].date; }).filter(function (d) { return d > today; }).sort();
    if (!up.length) { Logger.log('これからの予定がないため、見本を作れませんでした。アプリで予定を1つ追加してから試してください。'); return; }
    day = up[0];
  }
  const n = sendReminders(day, c, me);
  Logger.log(n ? me + ' にリマインドの見本(' + md(day) + 'の予定)を送りました。' : '見本を作れませんでした(その日の予定がすべて不参加など)。');
}
function testDigest() {
  const c = loadCtx(); const me = SENDER;
  const n = sendDigest(nowJst().date, c, me);
  Logger.log(n ? me + ' に週まとめの見本を送りました。' : '来週の予定もアウトプットもないため、週まとめは送られない状態です。');
}
function pickTestUid(c) { if (c.members[ADMIN_UID]) return ADMIN_UID; return Object.keys(c.members)[0] || ''; }

/* ---------------- 集計 ---------------- */
function stats(c, today) {
  const rec = Object.keys(c.events).map(function (k) { return c.events[k]; })
    .filter(function (e) { return e.date <= today && e.attended && e.attended.length; }).sort(byTime);
  const out = {};
  Object.keys(c.members).forEach(function (id) {
    let s = 0; for (var i = rec.length - 1; i >= 0; i--) { if (rec[i].attended.indexOf(id) >= 0) s++; else break; }
    out[id] = { streak: s, count: rec.filter(function (e) { return e.attended.indexOf(id) >= 0; }).length };
  });
  return out;
}
function byTime(a, b) { return (a.date + (a.start || '')).localeCompare(b.date + (b.start || '')); }
function nowJst() {
  const d = new Date();
  return { date: Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'), hour: Number(Utilities.formatDate(d, 'Asia/Tokyo', 'H')), dow: Number(Utilities.formatDate(d, 'Asia/Tokyo', 'u')) };
}
function addDays(date, n) { const p = date.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)).toISOString().slice(0, 10); }

/* ---------------- メール本文 ---------------- */
const BTN = 'display:inline-block;background:#C6F432;color:#0B1222;font-weight:700;text-decoration:none;padding:12px 20px;border-radius:12px';
const NOTE = 'font-size:13px;margin:0 0 10px;padding:10px 12px;border-radius:10px;';
function wrap(appName, lead, inner) {
  return '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;color:#0F172A">' +
    '<div style="background:#0B1222;color:#E8EEF9;padding:16px 20px;border-radius:14px 14px 0 0;font-weight:700">🔥 ' + h(appName) + '</div>' +
    '<div style="border:1px solid #D2DAE8;border-top:0;padding:20px;border-radius:0 0 14px 14px">' +
    '<p style="margin:0 0 14px">' + h(lead) + '</p>' + inner +
    '<p style="font-size:12px;color:#5B6782;margin:18px 0 0">アプリ:<a href="' + APP_URL + '">' + APP_URL + '</a><br>メールが不要なときは、アプリの設定で「受け取らない」にできます。</p>' +
    '</div></div>';
}
function eventBlock(e, members, withMemo) {
  const host = members[e.hostId] ? members[e.hostId].name : '';
  const row = function (k, v) { return v ? '<tr><td style="color:#5B6782;padding:4px 14px 4px 0;white-space:nowrap;vertical-align:top">' + k + '</td><td style="padding:4px 0">' + v + '</td></tr>' : ''; };
  return '<p style="font-size:18px;font-weight:700;margin:0 0 8px">' + h(title(e)) + '</p>' +
    '<table style="font-size:14px;border-collapse:collapse;margin-bottom:14px">' +
    row('種類', TYPES[e.type] || '') + row('日時', md(e.date) + ' ' + h(timeRange(e))) +
    row('場所', h(e.place || '') + (e.map ? ' <a href="' + h(e.map) + '">地図</a>' : '')) +
    row('担当', h(host)) + (withMemo ? row('メモ', h(e.memo || '').replace(/\n/g, '<br>')) : '') + '</table>';
}
function stat(big, label) {
  return '<td style="background:#F2F4F8;border-radius:12px;padding:12px 6px;text-align:center;width:33%"><div style="font-size:20px;font-weight:700">' + big + '</div><div style="font-size:11px;color:#5B6782">' + label + '</div></td>';
}

/* ---------------- 予定のフォーマット ---------------- */
function title(e) {
  if (e.type === 'grammar' && e.unit) { const u = UNITS[e.unit] || ''; return u + (e.title && e.title !== u ? ' — ' + e.title : ''); }
  return e.title || '';
}
function md(date) { const p = String(date).split('-').map(Number); const d = new Date(p[0], p[1] - 1, p[2]); return (p[1]) + '/' + p[2] + '(' + WD[d.getDay()] + ')'; }
function timeRange(e) { return e.start ? e.start + (e.end ? '〜' + e.end : '') : '時間未定'; }
function addHour(t) { const p = t.split(':').map(Number); return ('0' + Math.min(23, p[0] + 1)).slice(-2) + ':' + ('0' + p[1]).slice(-2); }
function nextDay(date) { const p = date.split('-').map(Number); const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + 1)); return d.toISOString().slice(0, 10).replace(/-/g, ''); }
function details(e) { return [TYPES[e.type] || '', e.memo || '', e.map || '', APP_URL].filter(String).join('\n'); }
function gcalUrl(e, appName) {
  const d = e.date.replace(/-/g, '');
  const dates = e.start ? d + 'T' + e.start.replace(':', '') + '00/' + d + 'T' + (e.end || addHour(e.start)).replace(':', '') + '00' : d + '/' + nextDay(e.date);
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=' + encodeURIComponent(appName + ':' + title(e)) +
    '&dates=' + dates + '&ctz=Asia%2FTokyo&details=' + encodeURIComponent(details(e)) + '&location=' + encodeURIComponent(e.place || '');
}
function ics(e, appName) {
  const stamp = function (d) { return d.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, ''); };
  const utc = function (date, time) { const a = date.split('-').map(Number), b = time.split(':').map(Number); return stamp(new Date(Date.UTC(a[0], a[1] - 1, a[2], b[0] - 9, b[1]))); };
  const esc = function (x) { return String(x || '').replace(/\\/g, '\\\\').replace(/\n/g, '\\n').replace(/,/g, '\\,').replace(/;/g, '\\;'); };
  const ds = e.start ? 'DTSTART:' + utc(e.date, e.start) : 'DTSTART;VALUE=DATE:' + e.date.replace(/-/g, '');
  const de = e.start ? 'DTEND:' + utc(e.date, e.end || addHour(e.start)) : 'DTEND;VALUE=DATE:' + nextDay(e.date);
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//study-board//JA', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'BEGIN:VEVENT',
    'UID:' + e.id + '@study-board', 'DTSTAMP:' + stamp(new Date()), ds, de, 'SUMMARY:' + esc(appName + ':' + title(e)),
    'LOCATION:' + esc(e.place), 'DESCRIPTION:' + esc(details(e)), 'URL:' + APP_URL,
    'BEGIN:VALARM', 'TRIGGER:-PT1H', 'ACTION:DISPLAY', 'DESCRIPTION:' + esc('1時間後に勉強会があります'), 'END:VALARM',
    'END:VEVENT', 'END:VCALENDAR'].join('\r\n');
}
function h(s) { return String(s || '').replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

/* ---------------- Firestore REST ---------------- */
function headers() { return { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'x-goog-user-project': PROJECT }; }
function fsList(col) {
  const out = []; let token = '';
  do {
    const res = UrlFetchApp.fetch(BASE + col + '?pageSize=300' + (token ? '&pageToken=' + encodeURIComponent(token) : ''), { headers: headers(), muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error(col + ' を読めません: ' + res.getContentText());
    const j = JSON.parse(res.getContentText());
    (j.documents || []).forEach(function (d) { const o = decode(d.fields || {}); o.id = d.name.split('/').pop(); out.push(o); });
    token = j.nextPageToken || '';
  } while (token);
  return out;
}
function fsGet(path) {
  const res = UrlFetchApp.fetch(BASE + path, { headers: headers(), muteHttpExceptions: true });
  if (res.getResponseCode() === 404) return null;
  if (res.getResponseCode() !== 200) throw new Error(path + ' を読めません: ' + res.getContentText());
  return decode(JSON.parse(res.getContentText()).fields || {});
}
function fsSet(path, obj) {
  const fields = {}; Object.keys(obj).forEach(function (k) { fields[k] = { stringValue: String(obj[k]) }; });
  const res = UrlFetchApp.fetch(BASE + path, { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ fields: fields }), headers: headers(), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error(path + ' に書けません: ' + res.getContentText());
}
function decode(f) { const o = {}; Object.keys(f).forEach(function (k) { o[k] = val(f[k]); }); return o; }
function val(v) {
  if ('stringValue' in v) return v.stringValue;
  if ('integerValue' in v) return Number(v.integerValue);
  if ('doubleValue' in v) return v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue;
  if ('nullValue' in v) return null;
  if ('timestampValue' in v) return v.timestampValue;
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(val);
  if ('mapValue' in v) return decode(v.mapValue.fields || {});
  return null;
}
