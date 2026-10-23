/**
 * =====================================================================
 *  ICTD Meeting Minutes — Google Apps Script Web App
 *  نموذج محضر اجتماع — الإدارة التنفيذية للاتصالات وتقنية المعلومات
 *  جامعة سليمان الراجحي — الإصدار 2.1
 *  Secured with: SRU Secure Access Kit v1.4 (@sr.edu.sa + AllowedUsers + AccessLog)
 * ---------------------------------------------------------------------
 *  Files:  Code.gs · Index.html (the form) · Logos.html · Denied.html · appsscript.json
 *
 *  Deploy: 1. Run `setup` once              → tabs Minutes · Decisions
 *          2. Run `setupAccessList` once    → tab AllowedUsers (owner + ICTD team)
 *          3. Run `installReminderTrigger`  → daily reminders 07:00 Riyadh
 *          4. Deploy ▸ Manage deployments ▸ ✏️ ▸ Version: New version
 *             Execute as: Me · Who has access: Anyone within Sulaiman Al Rajhi University
 *
 *  Page calls (google.script.run) — every one starts with requireSruUser_():
 *    apiNext(year)      -> { ok, minuteNo }           next free minute number
 *    apiList()          -> { ok, rows }               minute index (newest first)
 *    apiGet(no)         -> { ok, data }               one full minute (JSON)
 *    apiOpen()          -> { ok, rows }               open decisions across all minutes
 *    apiSave(payload)   -> { ok, minuteNo, created }  upsert minute + rebuild its decisions
 *    apiSend(minuteNo)  -> { ok, sent, recipients }   email the minute (PDF attached)
 *
 *  v2.1: the anonymous JSON API (doGet ?action=… / doPost) was removed —
 *        the data is reachable only through the signed-in page.
 * =====================================================================
 */

/* ▼▼▼ CONFIG — platform ▼▼▼ */
var APP_NAME        = 'محضر اجتماع';
var APP_NAME_EN     = 'ICTD Meeting Minutes';
var APP_VERSION     = '2.1';
var FAVICON_URL     = 'https://ictsru.github.io/ITOC/assets/itoc-logo.png';

var SHEET_ID        = '1pw9XIBc7Zt6uF3wJO_gOZQ_meAM9pCqKMMDWsSHEjZ4';  // ICTD Meeting Minutes - Data
var MINUTES_SHEET   = 'Minutes';
var DECISIONS_SHEET = 'Decisions';
var PREFIX          = 'ICTD-MOM';

var SENDER_NAME     = 'الإدارة التنفيذية للاتصالات وتقنية المعلومات';
var ADMIN_EMAIL     = 'm.elmahdy@sr.edu.sa';  // نسخة إدارية + ملخص التذكيرات اليومي
var REMIND_DAYS     = 3;                       // يبدأ التذكير قبل الاستحقاق بهذا العدد من الأيام
var DONE_STATUSES   = ['منجز', 'ملغي', 'مغلق'];
var TZ              = 'Asia/Riyadh';

/* ▼▼▼ CONFIG — SRU Secure Access Kit v1.4 ▼▼▼ */
var ALLOWED_DOMAIN        = 'sr.edu.sa';
/* Access log */
var LOG_ENABLED           = true;
var LOG_SPREADSHEET_ID    = SHEET_ID;          // same data file
var SHEET_LOG             = 'AccessLog';
var LOG_TIMEZONE          = TZ;
var LOG_DEDUP_MINUTES     = 30;                // same user within N min = one entry (0 = every open)
var LOG_RETENTION_DAYS    = 180;
/* Access list */
var ACCESS_MODE           = 'LIST';            // 'LIST' = only active emails in AllowedUsers · 'DOMAIN' = any @sr.edu.sa
var ACCESS_SPREADSHEET_ID = SHEET_ID;
var SHEET_ACCESS          = 'AllowedUsers';
var OWNER_EMAILS          = ['m.elmahdy@sr.edu.sa'];   // always allowed — the owner can never be locked out
var DEFAULT_ALLOWED = [
  // [email, name, role, notes] — trim to the people who prepare minutes before the first setupAccessList
  ['nabounar@sr.edu.sa',    'Nasser Abounar',  'admin', 'CIO'],
  ['m.elmahdy@sr.edu.sa',   'Mohamed ElMahdy', 'admin', 'Owner — IT Operations Manager'],
  ['n.nawaz@sr.edu.sa',     '', 'user', ''],
  ['m.omar@sr.edu.sa',      '', 'user', ''],
  ['m.saleh@sr.edu.sa',     '', 'user', ''],
  ['a.alrowedan@sr.edu.sa', '', 'user', ''],
  ['n.alrashidi@sr.edu.sa', '', 'user', ''],
  ['a.alawaji@sr.edu.sa',   '', 'user', ''],
  ['a.alaqeel@sr.edu.sa',   '', 'user', ''],
  ['a.alolayan@sr.edu.sa',  '', 'user', ''],
  ['s.albusayli@sr.edu.sa', '', 'user', ''],
  ['i.almutairi@sr.edu.sa', '', 'user', ''],
  ['a.elnadi@sr.edu.sa',    '', 'user', '']
];
var ACCESS_CACHE_SECONDS  = 120;               // list changes take effect within 2 minutes
/* ▲▲▲ CONFIG ▲▲▲ */

/* ============================== SCHEMA =============================== */
// SavedBy (v2.1) is appended AFTER PayloadJSON so rows saved by earlier versions keep their layout.
var MINUTE_HEADERS = ['Timestamp','MinuteNo','Title','Type','Place','DateG','DateH','Day','From','To',
  'Host','Chair','Secretary','Classification','Attendees','Apologies','PreviousFollowUp','Agenda',
  'Decisions','Notes','NextMeeting','PayloadJSON','SavedBy'];
var PAYLOAD_COL = MINUTE_HEADERS.indexOf('PayloadJSON') + 1;   // 22
var DECISION_HEADERS = ['MinuteNo','DateG','Seq','Decision','Owner','OwnerEmail','DueDate','Status',
  'MeetingTitle','UpdatedAt','LastReminded'];
var LOG_HEADERS    = ['timestamp', 'date', 'time', 'day', 'email', 'result', 'version'];
var ACCESS_HEADERS = ['email', 'name', 'role', 'active', 'notes'];
var DAYS_AR_       = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

/* ============================== WEB APP ============================== */

function doGet(e) {
  var email = currentEmail_();
  var status = accessStatus_(email);
  if (status !== 'ok') {
    logAccess_(email || '(unknown)', status === 'list' ? 'not-listed' : 'denied');
    var d = HtmlService.createTemplateFromFile('Denied');
    d.appName = APP_NAME; d.appNameEn = APP_NAME_EN;
    d.userEmail = email || '';
    d.reason = status;
    return d.evaluate().setTitle(APP_NAME + ' | غير مصرح')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  logAccess_(email, 'granted');

  var p = (e && e.parameter) || {};
  var t = HtmlService.createTemplateFromFile('Index');
  t.appName   = APP_NAME;
  t.appNameEn = APP_NAME_EN;
  t.version   = APP_VERSION;
  t.userEmail = email;
  t.dataJson  = safeJson_({
    initialNo: /^[A-Za-z0-9\-]{1,40}$/.test(String(p.no || '')) ? String(p.no) : ''   // deep link ?no=ICTD-MOM-2026-001
  });
  var out = t.evaluate().setTitle(APP_NAME_EN + ' | ' + APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}

/* ------------------- Server calls from the page --------------------- */
/*  EVERY function called via google.script.run starts with requireSruUser_(). */

function apiNext(year) {
  requireSruUser_();
  return { ok: true, minuteNo: nextNo(year) };
}

function apiList() {
  requireSruUser_();
  var sh = minutesSheet(), last = sh.getLastRow();
  if (last < 2) return { ok: true, rows: [] };
  var vals = sh.getRange(2, 1, last - 1, 14).getValues();
  var rows = vals
    .filter(function (v) { return String(v[1]).trim() !== ''; })
    .map(function (v) {
      return {
        minuteNo:  String(v[1]).trim(),
        title:     String(v[2]), type: String(v[3]), place: String(v[4]),
        dateG:     fmt(v[5], 'yyyy-MM-dd'),
        chair:     String(v[11]), secretary: String(v[12]),
        updated:   fmt(v[0], 'yyyy-MM-dd HH:mm'),
        timestamp: fmt(v[0], "yyyy-MM-dd'T'HH:mm:ss")
      };
    });
  rows.sort(function (a, b) { return a.minuteNo < b.minuteNo ? 1 : a.minuteNo > b.minuteNo ? -1 : 0; });
  return { ok: true, rows: rows };
}

function apiGet(no) {
  requireSruUser_();
  var d = loadPayload(no);
  return d ? { ok: true, data: d } : { ok: false, error: 'not_found' };
}

function apiOpen() {
  requireSruUser_();
  return { ok: true, rows: openDecisions().map(function (x) {
    x.due = fmt(x.due, 'yyyy-MM-dd'); return x;      // Dates can't cross google.script.run
  }) };
}

function apiSave(payload) {
  var email = requireSruUser_();
  return withLock_(function () { return saveMinute(payload || {}, email); });
}

function apiSend(minuteNo) {
  var email = requireSruUser_();
  return withLock_(function () { return sendMinute(minuteNo, email); });
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    return fn();
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* =============================== UTILS =============================== */
function ss() { return SpreadsheetApp.openById(SHEET_ID); }
function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}
function isEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '').trim()); }
function pad3(n) { n = String(n); while (n.length < 3) n = '0' + n; return n; }
function makeNo(year, seq) { return PREFIX + '-' + year + '-' + pad3(seq); }
function fmt(v, pattern) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, pattern);
  return String(v == null ? '' : v).trim();
}
function isDone(status) { return DONE_STATUSES.indexOf(String(status || '').trim()) > -1; }

function sheetFor(name, headers) {
  var book = ss();
  var sh = book.getSheetByName(name);
  if (!sh) {
    if (name === MINUTES_SHEET && book.getSheets().length === 1 && book.getSheets()[0].getLastRow() <= 1) {
      sh = book.getSheets()[0];
      sh.setName(name);
    } else {
      sh = book.insertSheet(name);
    }
  }
  var head = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (head.join('|') !== headers.join('|')) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#f0ebf8').setFontColor('#3a1464');
    sh.setFrozenRows(1);
  }
  return sh;
}
function minutesSheet()   { return sheetFor(MINUTES_SHEET,   MINUTE_HEADERS); }
function decisionsSheet() { return sheetFor(DECISIONS_SHEET, DECISION_HEADERS); }

function maxSeq(year) {
  var sh = minutesSheet(), last = sh.getLastRow();
  if (last < 2) return 0;
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  var re = new RegExp('^' + PREFIX + '-' + year + '-(\\d+)$'), max = 0;
  for (var i = 0; i < col.length; i++) {
    var m = re.exec(String(col[i][0]).trim());
    if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
  }
  return max;
}
function nextNo(year) {
  year = String(year || new Date().getFullYear());
  if (!/^\d{4}$/.test(year)) year = String(new Date().getFullYear());
  return makeNo(year, maxSeq(year) + 1);
}
function findRow(minuteNo) {
  var sh = minutesSheet(), last = sh.getLastRow();
  if (last < 2) return 0;
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() === String(minuteNo).trim()) return i + 2;
  }
  return 0;
}
function loadPayload(minuteNo) {
  var r = findRow(minuteNo);
  if (!r) return null;
  var raw = minutesSheet().getRange(r, PAYLOAD_COL, 1, 1).getValue();
  return raw ? JSON.parse(raw) : null;
}
function flatten(rows) {
  if (!rows || !rows.length) return '';
  return rows
    .filter(function (r) { return r.join('').trim() !== ''; })
    .map(function (r, i) { return (i + 1) + '. ' + r.filter(String).join(' | '); })
    .join('\n');
}

/* =============================== SAVE ================================ */
function saveMinute(data, savedBy) {
  var f = data.fields || {}, t = data.tables || {};
  var year = String(f.dateG || '').slice(0, 4);
  if (!/^\d{4}$/.test(year)) year = String(new Date().getFullYear());

  var minuteNo = String(f.docNo || '').trim() || nextNo(year);

  var row = [
    new Date(), minuteNo, f.title || '', f.type || '', f.place || '',
    f.dateG || '', f.dateH || '', f.day || '', f.from || '', f.to || '',
    f.host || '', f.chair || '', f.secretary || '', f.classification || '',
    flatten(t.attendees), flatten(t.apologies), flatten(t.previous),
    flatten(t.agenda), flatten(t.decisions), f.notes || '',
    [f.nextDate || '', f.nextTime || '', f.nextPlace || ''].filter(String).join(' — '),
    JSON.stringify(data),
    savedBy || ''
  ];

  var sh = minutesSheet(), r = findRow(minuteNo), created = false;
  if (r) { sh.getRange(r, 1, 1, row.length).setValues([row]); }
  else { sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]); created = true; }

  syncDecisions(minuteNo, f, t.decisions || [], t.attendees || []);
  return { ok: true, minuteNo: minuteNo, created: created };
}

/** email of an attendee whose name matches the decision owner */
function ownerEmail(name, attendees) {
  var n = String(name || '').trim();
  if (!n) return '';
  for (var i = 0; i < attendees.length; i++) {
    var a = attendees[i] || [];
    var an = String(a[0] || '').trim();
    if (!an) continue;
    if (an === n || an.indexOf(n) > -1 || n.indexOf(an) > -1) {
      if (isEmail(a[2])) return String(a[2]).trim();
    }
  }
  return '';
}

function syncDecisions(minuteNo, f, rows, attendees) {
  var sh = decisionsSheet(), last = sh.getLastRow();
  if (last >= 2) {
    var col = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = col.length - 1; i >= 0; i--) {
      if (String(col[i][0]).trim() === minuteNo) sh.deleteRow(i + 2);
    }
  }
  var out = [];
  (rows || []).forEach(function (r, i) {
    if (!r || String(r[0] || '').trim() === '') return;
    out.push([minuteNo, f.dateG || '', i + 1, r[0] || '', r[1] || '',
              ownerEmail(r[1], attendees), r[2] || '', r[3] || '',
              f.title || '', new Date(), '']);
  });
  if (out.length) {
    sh.getRange(sh.getLastRow() + 1, 1, out.length, DECISION_HEADERS.length).setValues(out);
  }
}

/* ========================= EMAIL — THE MINUTE ======================== */
var MAIL_CSS =
  "body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;direction:rtl;color:#241436;font-size:14px;line-height:1.8;margin:0;background:#f7f5fb}" +
  ".wrap{max-width:720px;margin:0 auto;background:#fff;border:1px solid #e2d8ef}" +
  ".bar{background:#3a1464;color:#fff;padding:16px 22px}" +
  ".bar .u{font-size:12px;color:#c9b7e6;letter-spacing:.04em}" +
  ".bar h1{margin:4px 0 0;font-size:18px;font-weight:700}" +
  ".body{padding:20px 22px}" +
  "h2{font-size:13px;background:#501e8c;color:#fff;padding:5px 10px;margin:20px 0 8px;border-radius:3px}" +
  "table{width:100%;border-collapse:collapse;margin-bottom:6px}" +
  "th,td{border:1px solid #cdbfe0;padding:6px 8px;text-align:right;font-size:12.5px;vertical-align:top}" +
  "th{background:#f0ebf8;color:#3a1464;font-weight:700}" +
  ".meta{width:100%;border:none;margin-bottom:4px}" +
  ".meta td{border:none;padding:2px 0;font-size:13px}" +
  ".meta .k{color:#6b5a80;width:130px}" +
  ".foot{padding:14px 22px;border-top:1px solid #e2d8ef;color:#8f83a2;font-size:11.5px}";

function minuteHtml(d) {
  var f = d.fields || {}, t = d.tables || {};
  function tbl(title, headers, rows, cols) {
    var body = (rows || []).filter(function (r) { return r.join('').trim() !== ''; });
    if (!body.length) return '';
    var h = '<h2>' + esc(title) + '</h2><table><tr><th style="width:34px">م</th>'
      + headers.map(function (x) { return '<th>' + esc(x) + '</th>'; }).join('') + '</tr>';
    body.forEach(function (r, i) {
      h += '<tr><td>' + (i + 1) + '</td>'
        + cols.map(function (c) { return '<td>' + esc(r[c]) + '</td>'; }).join('') + '</tr>';
    });
    return h + '</table>';
  }
  var html = '<html><head><meta charset="utf-8"><style>' + MAIL_CSS + '</style></head><body><div class="wrap">'
    + '<div class="bar"><div class="u">جامعة سليمان الراجحي · الإدارة التنفيذية للاتصالات وتقنية المعلومات</div>'
    + '<h1>محضر اجتماع — ' + esc(f.title || '') + '</h1></div><div class="body">'
    + '<table class="meta">'
    + '<tr><td class="k">رقم المحضر</td><td><b>' + esc(f.docNo) + '</b></td></tr>'
    + '<tr><td class="k">التاريخ</td><td>' + esc(f.dateG) + ' — من ' + esc(f.from) + ' إلى ' + esc(f.to) + '</td></tr>'
    + '<tr><td class="k">المكان</td><td>' + esc(f.place) + '</td></tr>'
    + '<tr><td class="k">رئيس الاجتماع</td><td>' + esc(f.chair) + '</td></tr>'
    + '<tr><td class="k">أمين السر</td><td>' + esc(f.secretary) + '</td></tr>'
    + '<tr><td class="k">التصنيف</td><td>' + esc(f.classification) + '</td></tr>'
    + '</table>'
    + tbl('الحضور', ['الاسم','الجهة / الوحدة','البريد الإلكتروني','رقم الجوال'], t.attendees, [0,1,2,3])
    + tbl('المعتذرون', ['الاسم','الجهة / الوحدة','سبب الاعتذار','من ينوب عنه'], t.apologies, [0,1,2,3])
    + tbl('متابعة قرارات الاجتماع السابق', ['القرار','المسؤول','الحالة','ملاحظات'], t.previous, [0,1,2,3])
    + tbl('جدول الأعمال والمناقشات', ['البند','ملخص المناقشة'], t.agenda, [0,1])
    + tbl('القرارات ومهام المتابعة', ['القرار / المهمة','المسؤول','تاريخ الاستحقاق','الحالة'], t.decisions, [0,1,2,3]);

  if (String(f.notes || '').trim()) {
    html += '<h2>التوصيات والملاحظات الختامية</h2><div>' + esc(f.notes) + '</div>';
  }
  var next = [f.nextDate, f.nextTime, f.nextPlace].filter(String).join(' — ');
  if (next) html += '<h2>الاجتماع القادم</h2><div>' + esc(next) + '</div>';

  html += '</div><div class="foot">نموذج ICTD-FRM-MOM-01 · أُرسل آلياً من نظام محاضر اجتماعات '
    + esc(SENDER_NAME) + '</div></div></body></html>';
  return html;
}

function sendMinute(minuteNo, sentBy) {
  var d = loadPayload(minuteNo);
  if (!d) return { ok: false, error: 'not_found' };

  var f = d.fields || {}, t = d.tables || {};
  var to = [];
  (t.attendees || []).forEach(function (r) { if (isEmail(r[2])) to.push(String(r[2]).trim()); });
  if (isEmail(ADMIN_EMAIL)) to.push(ADMIN_EMAIL);

  var seen = {}, list = [];
  to.forEach(function (a) { var k = a.toLowerCase(); if (!seen[k]) { seen[k] = 1; list.push(a); } });
  if (!list.length) return { ok: false, error: 'no_recipients' };

  var html = minuteHtml(d);
  var name = 'محضر-' + minuteNo;
  var pdf;
  try {
    pdf = Utilities.newBlob(html, 'text/html', name + '.html')
      .getAs('application/pdf').setName(name + '.pdf');
  } catch (err) { pdf = null; }   // البريد يُرسل بالنص المنسّق حتى لو تعذّر توليد PDF

  MailApp.sendEmail({
    to: list.join(','),
    subject: 'محضر اجتماع ' + minuteNo + ' — ' + (f.title || ''),
    htmlBody: html,
    name: SENDER_NAME,
    replyTo: isEmail(sentBy) ? sentBy : '',
    attachments: pdf ? [pdf] : []
  });
  return { ok: true, sent: list.length, recipients: list, pdf: !!pdf };
}

/* ====================== REMINDERS — DUE DECISIONS ==================== */
function toDate(v) {
  if (v instanceof Date) return v;
  var s = String(v || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  var p = s.split('-');
  return new Date(+p[0], +p[1] - 1, +p[2]);
}
function daysBetween(a, b) {
  return Math.round((a.getTime() - b.getTime()) / 86400000);
}

/** every open decision */
function openDecisions() {
  var sh = decisionsSheet(), last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, DECISION_HEADERS.length).getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var out = [];
  vals.forEach(function (v, i) {
    if (isDone(v[7])) return;
    var due = toDate(v[6]);
    out.push({
      row: i + 2, minuteNo: v[0], seq: v[2], decision: v[3], owner: v[4],
      ownerEmail: v[5], due: v[6], status: v[7], meeting: v[8],
      daysLeft: due ? daysBetween(due, today) : null
    });
  });
  return out;
}

/** run daily (trigger) — reminds each owner, then sends the admin a digest */
function sendReminders() {
  var items = openDecisions().filter(function (x) {
    return x.daysLeft !== null && x.daysLeft <= REMIND_DAYS;
  });
  if (!items.length) return;

  var sh = decisionsSheet();
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var byOwner = {};

  items.forEach(function (x) {
    if (!isEmail(x.ownerEmail)) return;
    (byOwner[x.ownerEmail] = byOwner[x.ownerEmail] || []).push(x);
  });

  Object.keys(byOwner).forEach(function (email) {
    var rows = byOwner[email];
    MailApp.sendEmail({
      to: email,
      subject: 'تذكير: ' + rows.length + ' مهمة متابعة مستحقة — محاضر اجتماعات ICTD',
      htmlBody: remindHtml(rows[0].owner, rows, false),
      name: SENDER_NAME
    });
    rows.forEach(function (x) { sh.getRange(x.row, 11).setValue(stamp); });
  });

  if (isEmail(ADMIN_EMAIL)) {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: 'ملخص متابعة القرارات — ' + items.length + ' مهمة مستحقة أو متأخرة',
      htmlBody: remindHtml('', items, true),
      name: SENDER_NAME
    });
  }
}

function remindHtml(owner, rows, isDigest) {
  var head = isDigest
    ? '<p>هذا ملخص جميع مهام المتابعة المستحقة أو المتأخرة عبر محاضر الاجتماعات:</p>'
    : '<p>السلام عليكم' + (owner ? ' أ. ' + esc(owner) : '') +
      '، فيما يلي مهام المتابعة المسندة إليكم والمستحقة خلال ' + REMIND_DAYS + ' أيام أو المتأخرة:</p>';

  var h = '<html><head><meta charset="utf-8"><style>' + MAIL_CSS + '</style></head><body><div class="wrap">'
    + '<div class="bar"><div class="u">جامعة سليمان الراجحي · ' + esc(SENDER_NAME) + '</div>'
    + '<h1>' + (isDigest ? 'ملخص متابعة القرارات' : 'تذكير بمهام المتابعة') + '</h1></div>'
    + '<div class="body">' + head
    + '<table><tr><th>القرار / المهمة</th>'
    + (isDigest ? '<th>المسؤول</th>' : '')
    + '<th>الاجتماع</th><th>الاستحقاق</th><th>الحالة</th><th>المتبقي</th></tr>';

  rows.sort(function (a, b) { return (a.daysLeft || 0) - (b.daysLeft || 0); }).forEach(function (x) {
    var left = x.daysLeft < 0 ? ('<b style="color:#9d1f18">متأخر ' + Math.abs(x.daysLeft) + ' يوم</b>')
             : x.daysLeft === 0 ? '<b style="color:#8a5b06">اليوم</b>'
             : (x.daysLeft + ' يوم');
    h += '<tr><td>' + esc(x.decision) + '</td>'
      + (isDigest ? '<td>' + esc(x.owner) + '</td>' : '')
      + '<td>' + esc(x.minuteNo) + '</td><td>' + esc(fmt(x.due, 'yyyy-MM-dd')) + '</td>'
      + '<td>' + esc(x.status) + '</td><td>' + left + '</td></tr>';
  });

  return h + '</table></div><div class="foot">رسالة آلية من نظام محاضر اجتماعات ICTD — '
    + 'يرجى تحديث حالة المهمة في محضر الاجتماع المعني.</div></div></body></html>';
}

/* ============================== ACCESS =============================== */

function currentEmail_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}
function isAllowed_(email) {
  return !!email && email.endsWith('@' + ALLOWED_DOMAIN);
}
function requireSruUser_() {
  var email = currentEmail_();
  var status = accessStatus_(email);
  if (status === 'domain') throw new Error('غير مصرح — الدخول لمنسوبي الجامعة فقط');
  if (status === 'list') throw new Error('غير مصرح — حسابك غير مضاف إلى قائمة المصرّح لهم');
  return email;
}

/** 'ok' | 'domain' (not @sr.edu.sa) | 'list' (not on the active list) */
function accessStatus_(email) {
  if (!isAllowed_(email)) return 'domain';
  if (ACCESS_MODE !== 'LIST') return 'ok';
  if (OWNER_EMAILS.map(function (e) { return e.toLowerCase(); }).indexOf(email) !== -1) return 'ok';
  return allowedEmails_().indexOf(email) !== -1 ? 'ok' : 'list';
}

/** Active emails from the AllowedUsers tab (cached briefly). */
function allowedEmails_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('acl_v1');
  if (hit) return JSON.parse(hit);
  var list = [];
  try {
    var sh = accessSheet_(false);
    if (sh && sh.getLastRow() > 1) {
      list = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
        .filter(function (r) { return r[0] && r[3] !== false && String(r[3]).toUpperCase() !== 'FALSE'; })
        .map(function (r) { return String(r[0]).trim().toLowerCase(); });
    }
  } catch (e) { console.error('allowedEmails_ failed: ' + e); }
  cache.put('acl_v1', JSON.stringify(list), ACCESS_CACHE_SECONDS);
  return list;
}

function accessSheet_(create) {
  var book = ACCESS_SPREADSHEET_ID ? SpreadsheetApp.openById(ACCESS_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = book.getSheetByName(SHEET_ACCESS);
  if (!sh && create) {
    sh = book.insertSheet(SHEET_ACCESS);
    sh.getRange(1, 1, 1, ACCESS_HEADERS.length).setValues([ACCESS_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#501e8c');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    [260, 200, 100, 70, 260].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(2, 4, 300, 1).insertCheckboxes();
    sh.getRange('A1').setNote('البريد الجامعي كاملًا — name@sr.edu.sa');
    sh.getRange('C1').setNote('admin / user (للتوثيق — يمكن استخدامه لاحقًا لصلاحيات أدق)');
    sh.getRange('D1').setNote('ألغِ العلامة لإيقاف الوصول دون حذف الصف');
  }
  return sh;
}

/** Run once: creates the AllowedUsers tab and adds the owner + default list (existing rows kept). */
function setupAccessList() {
  var sh = accessSheet_(true);
  var n = Math.max(sh.getLastRow() - 1, 0);
  var data = n ? sh.getRange(2, 1, n, 5).getValues() : [];

  var kept = data.filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) { return [String(r[0]).trim(), r[1], r[2], r[3] === true || String(r[3]).toUpperCase() === 'TRUE', r[4]]; });
  var have = kept.map(function (r) { return r[0].toLowerCase(); });

  var seed = OWNER_EMAILS.map(function (e) { return [e, 'Owner', 'admin', 'owner']; }).concat(DEFAULT_ALLOWED);
  var added = 0;
  seed.forEach(function (r) {
    var e = String(r[0]).trim().toLowerCase();
    if (e && have.indexOf(e) === -1) {
      kept.push([e, r[1] || '', r[2] || 'user', true, r[3] || 'added by setupAccessList']);
      have.push(e); added++;
    }
  });

  if (n) sh.getRange(2, 1, n, 5).clearContent();
  var rows = Math.max(n, kept.length, 50);
  sh.getRange(2, 4, rows, 1).insertCheckboxes();
  if (kept.length) sh.getRange(2, 1, kept.length, 5).setValues(kept);

  CacheService.getScriptCache().remove('acl_v1');
  console.log('AllowedUsers ready: ' + kept.length + ' entries (' + added + ' added).');
  try { SpreadsheetApp.getActive().toast('قائمة المصرّح لهم جاهزة: ' + kept.length + ' بريدًا (أُضيف ' + added + ').', SHEET_ACCESS, 8); } catch (e) {}
}

/* ============================ ACCESS LOG ============================= */

/** Appends one row to AccessLog. Never blocks the page if logging fails. */
function logAccess_(email, result) {
  if (!LOG_ENABLED) return;
  try {
    if (LOG_DEDUP_MINUTES > 0 && result === 'granted') {
      var cache = CacheService.getScriptCache();
      var key = 'log_' + email;
      if (cache.get(key)) return;
      cache.put(key, '1', LOG_DEDUP_MINUTES * 60);
    }
    var now = new Date();
    var f = function (p) { return Utilities.formatDate(now, LOG_TIMEZONE, p); };
    getLogSheet_().appendRow([now, f('yyyy-MM-dd'), f('HH:mm:ss'), DAYS_AR_[Number(f('u')) % 7],
      email, result === 'granted' ? 'دخول' : (result === 'not-listed' ? 'غير مدرج' : 'مرفوض'), APP_VERSION]);
  } catch (e) {
    console.error('logAccess_ failed: ' + e);
  }
}

function getLogSheet_() {
  var book = LOG_SPREADSHEET_ID ? SpreadsheetApp.openById(LOG_SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('No spreadsheet: bind the script to a Sheet or set LOG_SPREADSHEET_ID');
  var sh = book.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = book.insertSheet(SHEET_LOG);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#3a1464');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    [170, 100, 80, 80, 260, 80, 70].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  }
  return sh;
}

/** Menu: quick summary of the last 30 days. */
function showLogSummary() {
  var sh = getLogSheet_();
  var ui = SpreadsheetApp.getUi();
  var last = sh.getLastRow();
  if (last < 2) { ui.alert('سجل الدخول', 'لا توجد سجلات بعد.', ui.ButtonSet.OK); return; }
  var since = new Date(Date.now() - 30 * 864e5);
  var rows = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues()
    .filter(function (r) { return r[0] instanceof Date && r[0] >= since && r[5] === 'دخول'; });
  var users = {};
  rows.forEach(function (r) { users[r[4]] = (users[r[4]] || 0) + 1; });
  var top = Object.keys(users).sort(function (a, b) { return users[b] - users[a]; }).slice(0, 10)
    .map(function (u, i) { return (i + 1) + '. ' + u + ' — ' + users[u]; }).join('\n');
  ui.alert('سجل الدخول — آخر 30 يومًا — ' + APP_NAME,
    'عدد الزيارات: ' + rows.length + '\nعدد المستخدمين: ' + Object.keys(users).length + '\n\nالأكثر دخولًا:\n' + (top || '—'),
    ui.ButtonSet.OK);
}

/** Menu: delete entries older than LOG_RETENTION_DAYS. */
function purgeOldLogs() {
  var sh = getLogSheet_();
  var last = sh.getLastRow();
  var removed = 0;
  if (last > 1) {
    var cutoff = new Date(Date.now() - LOG_RETENTION_DAYS * 864e5);
    var stamps = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = stamps.length - 1; i >= 0; i--) {
      if (stamps[i][0] instanceof Date && stamps[i][0] < cutoff) { sh.deleteRow(i + 2); removed++; }
    }
  }
  SpreadsheetApp.getActive().toast('حُذف ' + removed + ' سجل أقدم من ' + LOG_RETENTION_DAYS + ' يومًا.', APP_NAME_EN, 6);
}

/* ============================ SETUP / MENU =========================== */

function setup() {
  minutesSheet();
  decisionsSheet();
}

/** run once to schedule the daily reminder at 07:00 Riyadh time */
function installReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(7).everyDays(1).create();
}

/** optional helper — check what today's reminder run would send */
function previewReminders() {
  var items = openDecisions().filter(function (x) {
    return x.daysLeft !== null && x.daysLeft <= REMIND_DAYS;
  });
  Logger.log(JSON.stringify(items, null, 2));
}

/** Sheet menu (bound scripts) — one onOpen for the whole project. */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu(APP_NAME_EN)
      .addItem('إعداد تبويبات المحاضر (setup)', 'setup')
      .addItem('تفعيل التذكيرات اليومية', 'installReminderTrigger')
      .addSeparator()
      .addItem('إعداد قائمة المصرّح لهم (AllowedUsers)', 'setupAccessList')
      .addItem('ملخص سجل الدخول (30 يومًا)', 'showLogSummary')
      .addItem('حذف سجلات الدخول القديمة', 'purgeOldLogs')
      .addSeparator()
      .addItem('عرض رابط المنصة', 'showWebAppUrl')
      .addToUi();
  } catch (e) { /* standalone script: no sheet UI */ }
}

function showWebAppUrl() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('رابط ' + APP_NAME, ScriptApp.getService().getUrl() || 'لم يُنشر بعد: Deploy ▸ New deployment ▸ Web app', ui.ButtonSet.OK);
}

/* ============================== HELPERS ============================== */

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}
function safeJson_(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}
