/**
 * 英語勉強会ボード メール送信係(Google Apps Script)
 *
 * 1分ごとに Firestore を確認して、次のメールを送ります。
 *  ・新しい予定が追加されたとき → メンバー全員(メール通知オンの人)にお知らせ
 *  ・誰かが「参加」を押したとき → その人に予約確定メール(カレンダー用 .ics 付き)
 *
 * 初回だけ: 関数「setup」を選んで ▶実行 → Googleの許可画面で許可。
 * 送信元は、このスクリプトを動かしている Google アカウントの Gmail です。
 */
const PROJECT = 'study-board-9aefe';
const APP_URL = 'https://shintaroyade.github.io/study-board/';
const DEFAULT_NAME = '英語勉強会';
const DEFAULT_INVITE = 'OUTPUT2026';
const BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents/';
const TYPES = { grammar: '英文法', lesson: 'メンバー授業', other: 'その他' };
const UNITS = { bunkei:'文型', hinshi:'品詞', jisei:'時制', kanryo:'完了形', jodoushi:'助動詞', ukemi:'受動態', futeishi:'不定詞', doumeishi:'動名詞', bunshi:'分詞', hikaku:'比較', kankei:'関係詞', katei:'仮定法', wahou:'話法', meishi:'名詞・冠詞', zenchi:'前置詞', setsuzoku:'接続詞', tokushu:'特殊構文' };
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

/** 1分ごとに自動で動く */
function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const since = Number(props.getProperty('since') || Date.now());
    const settings = fsGet('settings/app');
    const appName = (settings && settings.name) || DEFAULT_NAME;
    const members = {}; fsList('members').forEach(function (m) { members[m.id] = m; });
    const contacts = {}; fsList('contacts').forEach(function (c) { contacts[c.id] = c; });
    const events = {}; fsList('events').forEach(function (e) { events[e.id] = e; });
    const wants = function (uid) { const c = contacts[uid]; return c && c.email && c.notify !== false && members[uid]; };

    // 1) 新しい予定 → 全員にお知らせ
    Object.keys(events).forEach(function (id) {
      const e = events[id]; const key = 'ev_' + id;
      if (props.getProperty(key) || !(Number(e.createdAt) >= since)) return;
      props.setProperty(key, String(Date.now()));
      const by = members[e.createdBy] ? members[e.createdBy].name : 'メンバー';
      Object.keys(contacts).forEach(function (uid) {
        if (!wants(uid)) return;
        const name = members[uid].name || '';
        MailApp.sendEmail({
          to: contacts[uid].email,
          subject: '【' + appName + '】新しい予定:' + md(e.date) + ' ' + title(e),
          name: appName,
          htmlBody: mailHtml(appName, name + 'さん、' + by + 'さんが新しい予定を追加しました。', e, members,
            '<a href="' + APP_URL + '" style="' + BTN + '">アプリで参加を返事する</a>'),
        });
      });
    });

    // 2) 「参加」→ 予約確定メール
    fsList('rsvps').forEach(function (r) {
      if (r.status !== 'yes') return;
      const key = 'rs_' + r.id;
      if (props.getProperty(key) || !(Number(r.at) >= since)) return;
      props.setProperty(key, String(Date.now()));
      const e = events[r.eventId];
      if (!e || !wants(r.uid)) return;
      MailApp.sendEmail({
        to: contacts[r.uid].email,
        subject: '【予約確定】' + md(e.date) + ' ' + title(e) + '|' + appName,
        name: appName,
        htmlBody: mailHtml(appName, (members[r.uid].name || '') + 'さん、参加の予約が確定しました。', e, members,
          '<a href="' + gcalUrl(e, appName) + '" style="' + BTN + '">Googleカレンダーに追加</a>' +
          '<p style="font-size:12px;color:#5B6782;margin:10px 0 0">iPhoneなどは、添付の「study.ics」を開くとカレンダーに追加できます。</p>'),
        attachments: [Utilities.newBlob(ics(e, appName), 'text/calendar', 'study.ics')],
      });
    });

    // 古い送信記録を掃除(60日)
    const cut = Date.now() - 60 * 864e5; const all = props.getProperties();
    Object.keys(all).forEach(function (k) { if (/^(ev|rs)_/.test(k) && Number(all[k]) < cut) props.deleteProperty(k); });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- メール本文 ---------------- */
const BTN = 'display:inline-block;background:#C6F432;color:#0B1222;font-weight:700;text-decoration:none;padding:12px 20px;border-radius:12px';
function mailHtml(appName, lead, e, members, actions) {
  const host = members[e.hostId] ? members[e.hostId].name : '';
  const row = function (k, v) { return v ? '<tr><td style="color:#5B6782;padding:4px 14px 4px 0;white-space:nowrap;vertical-align:top">' + k + '</td><td style="padding:4px 0">' + v + '</td></tr>' : ''; };
  return '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;color:#0F172A">' +
    '<div style="background:#0B1222;color:#E8EEF9;padding:16px 20px;border-radius:14px 14px 0 0;font-weight:700">🔥 ' + h(appName) + '</div>' +
    '<div style="border:1px solid #D2DAE8;border-top:0;padding:20px;border-radius:0 0 14px 14px">' +
    '<p style="margin:0 0 14px">' + h(lead) + '</p>' +
    '<p style="font-size:18px;font-weight:700;margin:0 0 8px">' + h(title(e)) + '</p>' +
    '<table style="font-size:14px;border-collapse:collapse;margin-bottom:16px">' +
    row('種類', TYPES[e.type] || '') + row('日時', md(e.date) + ' ' + h(timeRange(e))) +
    row('場所', h(e.place || '') + (e.map ? ' <a href="' + h(e.map) + '">地図</a>' : '')) +
    row('担当', h(host)) + row('メモ', h(e.memo || '').replace(/\n/g, '<br>')) +
    '</table>' + actions +
    '<p style="font-size:12px;color:#5B6782;margin:18px 0 0">アプリ:<a href="' + APP_URL + '">' + APP_URL + '</a><br>メールが不要なときは、アプリの設定で「受け取らない」にできます。</p>' +
    '</div></div>';
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
