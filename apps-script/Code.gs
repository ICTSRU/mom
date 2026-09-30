/**
 * ICTD Meeting Minutes — Google Apps Script backend
 * الإدارة التنفيذية للاتصالات وتقنية المعلومات — جامعة سليمان الراجحي
 * الإصدار 2.0
 *
 * Deploy: Extensions ▸ Apps Script ▸ paste this file ▸ Run `setup` once
 *         ▸ Deploy ▸ New deployment ▸ Web app
 *           Execute as: Me · Who has access: Anyone
 *         ▸ Run `installReminderTrigger` once to enable the daily reminders
 *
 * Endpoints
 *   GET  ?action=next&year=2026    -> { ok, minuteNo }   next free minute number
 *   GET  ?action=list              -> { ok, rows }       minute index (newest first, formatted dates)
 *   GET  ?action=get&no=ICTD-...   -> { ok, data }       one full minute (JSON)
 *   GET  ?action=open              -> { ok, rows }       open decisions across all minutes
 *   POST { ...form JSON }          -> { ok, minuteNo }   upsert minute + rebuild its decisions
 *   POST { action:'send', minuteNo}-> { ok, sent, recipients }  email the minute (PDF attached)
 */

/* =========================== CONFIGURATION =========================== */
var SHEET_ID        = '1pw9XIBc7Zt6uF3wJO_gOZQ_meAM9pCqKMMDWsSHEjZ4';
var MINUTES_SHEET   = 'Minutes';
var DECISIONS_SHEET = 'Decisions';
var PREFIX          = 'ICTD-MOM';

var SENDER_NAME     = 'الإدارة التنفيذية للاتصالات وتقنية المعلومات';
var ADMIN_EMAIL     = 'm.elmahdy@sr.edu.sa';  // نسخة إدارية + ملخص التذكيرات اليومي
var REMIND_DAYS     = 3;                       // يبدأ التذكير قبل الاستحقاق بهذا العدد من الأيام
var DONE_STATUSES   = ['منجز', 'ملغي', 'مغلق'];
var TZ              = 'Asia/Riyadh';

/* ============================== SCHEMA =============================== */
var MINUTE_HEADERS = ['Timestamp','MinuteNo','Title','Type','Place','DateG','DateH','Day','From','To',
  'Host','Chair','Secretary','Classification','Attendees','Apologies','PreviousFollowUp','Agenda',
  'Decisions','Notes','NextMeeting','PayloadJSON'];
var DECISION_HEADERS = ['MinuteNo','DateG','Seq','Decision','Owner','OwnerEmail','DueDate','Status',
  'MeetingTitle','UpdatedAt','LastReminded'];

/* =============================== UTILS =============================== */
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ss(){ return SpreadsheetApp.openById(SHEET_ID); }
function esc(v){
  return String(v == null ? '' : v)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>');
}
function isEmail(v){ return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '').trim()); }
function pad3(n){ n = String(n); while(n.length < 3) n = '0' + n; return n; }
function makeNo(year, seq){ return PREFIX + '-' + year + '-' + pad3(seq); }
function fmt(v, pattern){
  if(v instanceof Date) return Utilities.formatDate(v, TZ, pattern);
  return String(v == null ? '' : v).trim();
}
function isDone(status){ return DONE_STATUSES.indexOf(String(status || '').trim()) > -1; }

function sheetFor(name, headers){
  var book = ss();
  var sh = book.getSheetByName(name);
  if(!sh){
    if(name === MINUTES_SHEET && book.getSheets().length === 1 && book.getSheets()[0].getLastRow() <= 1){
      sh = book.getSheets()[0];
      sh.setName(name);
    } else {
      sh = book.insertSheet(name);
    }
  }
  var head = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if(head.join('|') !== headers.join('|')){
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#f0ebf8').setFontColor('#3a1464');
  return sh;
}
function minutesSheet(){   return sheetFor(MINUTES_SHEET,   MINUTE_HEADERS); }
function decisionsSheet(){ return sheetFor(DECISIONS_SHEET, DECISION_HEADERS); }

function maxSeq(year){
  var sh = minutesSheet(), last = sh.getLastRow();
  if(last < 2) return 0;
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  var re = new RegExp('^' + PREFIX + '-' + year + '-(\\d+)$'), max = 0;
  for(var i = 0; i < col.length; i++){
    var m = re.exec(String(col[i][0]).trim());
    if(m){ var n = parseInt(m[1], 10); if(n > max) max = n; }
  }
  return max;
}
function nextNo(year){
  year = String(year || new Date().getFullYear());
  if(!/^\d{4}$/.test(year)) year = String(new Date().getFullYear());
  return makeNo(year, maxSeq(year) + 1);
}
function findRow(minuteNo){
  var sh = minutesSheet(), last = sh.getLastRow();
  if(last < 2) return 0;
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  for(var i = 0; i < col.length; i++){
    if(String(col[i][0]).trim() === String(minuteNo).trim()) return i + 2;
  }
  return 0;
}
function loadPayload(minuteNo){
  var r = findRow(minuteNo);
  if(!r) return null;
  var raw = minutesSheet().getRange(r, MINUTE_HEADERS.length, 1, 1).getValue();
  return raw ? JSON.parse(raw) : null;
}
function flatten(rows){
  if(!rows || !rows.length) return '';
  return rows
    .filter(function(r){ return r.join('').trim() !== ''; })
    .map(function(r, i){ return (i + 1) + '. ' + r.filter(String).join(' | '); })
    .join('\n');
}

/* ================================ GET ================================ */
function doGet(e){
  var p = (e && e.parameter) || {};
  var action = String(p.action || '').toLowerCase();
  try{
    if(action === 'next') return json({ok:true, minuteNo: nextNo(p.year)});
    if(action === 'get'){
      var d = loadPayload(p.no);
      return d ? json({ok:true, data:d}) : json({ok:false, error:'not_found'});
    }
    if(action === 'list'){
      var sh = minutesSheet(), last = sh.getLastRow();
      if(last < 2) return json({ok:true, rows:[]});
      var vals = sh.getRange(2, 1, last - 1, 14).getValues();
      var rows = vals
        .filter(function(v){ return String(v[1]).trim() !== ''; })
        .map(function(v){
          return {
            minuteNo:  String(v[1]).trim(),
            title:     v[2], type: v[3], place: v[4],
            dateG:     fmt(v[5], 'yyyy-MM-dd'),
            chair:     v[11], secretary: v[12],
            updated:   fmt(v[0], 'yyyy-MM-dd HH:mm'),
            timestamp: fmt(v[0], "yyyy-MM-dd'T'HH:mm:ss")
          };
        });
      rows.sort(function(a, b){ return a.minuteNo < b.minuteNo ? 1 : a.minuteNo > b.minuteNo ? -1 : 0; });
      return json({ok:true, rows: rows});
    }
    if(action === 'open') return json({ok:true, rows: openDecisions()});
    return json({ok:true, service:'ICTD Meeting Minutes API', version:'2.0'});
  }catch(err){ return json({ok:false, error:String(err)}); }
}

/* =============================== POST ================================ */
function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(25000);
    var data = JSON.parse(e.postData.contents);

    if(String(data.action || '').toLowerCase() === 'send'){
      return json(sendMinute(data.minuteNo, data.extra));
    }
    return json(saveMinute(data));
  }catch(err){
    return json({ok:false, error:String(err)});
  }finally{
    try{ lock.releaseLock(); }catch(x){}
  }
}

function saveMinute(data){
  var f = data.fields || {}, t = data.tables || {};
  var year = String(f.dateG || '').slice(0,4);
  if(!/^\d{4}$/.test(year)) year = String(new Date().getFullYear());

  var minuteNo = String(f.docNo || '').trim() || nextNo(year);

  var row = [
    new Date(), minuteNo, f.title || '', f.type || '', f.place || '',
    f.dateG || '', f.dateH || '', f.day || '', f.from || '', f.to || '',
    f.host || '', f.chair || '', f.secretary || '', f.classification || '',
    flatten(t.attendees), flatten(t.apologies), flatten(t.previous),
    flatten(t.agenda), flatten(t.decisions), f.notes || '',
    [f.nextDate || '', f.nextTime || '', f.nextPlace || ''].filter(String).join(' — '),
    JSON.stringify(data)
  ];

  var sh = minutesSheet(), r = findRow(minuteNo), created = false;
  if(r){ sh.getRange(r, 1, 1, row.length).setValues([row]); }
  else { sh.appendRow(row); created = true; }

  syncDecisions(minuteNo, f, t.decisions || [], t.attendees || []);
  return {ok:true, minuteNo: minuteNo, created: created};
}

/** email of an attendee whose name matches the decision owner */
function ownerEmail(name, attendees){
  var n = String(name || '').trim();
  if(!n) return '';
  for(var i = 0; i < attendees.length; i++){
    var a = attendees[i] || [];
    var an = String(a[0] || '').trim();
    if(!an) continue;
    if(an === n || an.indexOf(n) > -1 || n.indexOf(an) > -1){
      if(isEmail(a[2])) return String(a[2]).trim();
    }
  }
  return '';
}

function syncDecisions(minuteNo, f, rows, attendees){
  var sh = decisionsSheet(), last = sh.getLastRow();
  if(last >= 2){
    var col = sh.getRange(2, 1, last - 1, 1).getValues();
    for(var i = col.length - 1; i >= 0; i--){
      if(String(col[i][0]).trim() === minuteNo) sh.deleteRow(i + 2);
    }
  }
  var out = [];
  (rows || []).forEach(function(r, i){
    if(!r || String(r[0] || '').trim() === '') return;
    out.push([minuteNo, f.dateG || '', i + 1, r[0] || '', r[1] || '',
              ownerEmail(r[1], attendees), r[2] || '', r[3] || '',
              f.title || '', new Date(), '']);
  });
  if(out.length){
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

function minuteHtml(d){
  var f = d.fields || {}, t = d.tables || {};
  function tbl(title, headers, rows, cols){
    var body = (rows || []).filter(function(r){ return r.join('').trim() !== ''; });
    if(!body.length) return '';
    var h = '<h2>' + esc(title) + '</h2><table><tr><th style="width:34px">م</th>'
      + headers.map(function(x){ return '<th>' + esc(x) + '</th>'; }).join('') + '</tr>';
    body.forEach(function(r, i){
      h += '<tr><td>' + (i + 1) + '</td>'
        + cols.map(function(c){ return '<td>' + esc(r[c]) + '</td>'; }).join('') + '</tr>';
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

  if(String(f.notes || '').trim()){
    html += '<h2>التوصيات والملاحظات الختامية</h2><div>' + esc(f.notes) + '</div>';
  }
  var next = [f.nextDate, f.nextTime, f.nextPlace].filter(String).join(' — ');
  if(next) html += '<h2>الاجتماع القادم</h2><div>' + esc(next) + '</div>';

  html += '</div><div class="foot">نموذج ICTD-FRM-MOM-01 · أُرسل آلياً من نظام محاضر اجتماعات '
    + esc(SENDER_NAME) + '</div></div></body></html>';
  return html;
}

function sendMinute(minuteNo, extra){
  var d = loadPayload(minuteNo);
  if(!d) return {ok:false, error:'not_found'};

  var f = d.fields || {}, t = d.tables || {};
  var to = [];
  (t.attendees || []).forEach(function(r){ if(isEmail(r[2])) to.push(String(r[2]).trim()); });
  (extra || []).forEach(function(x){ if(isEmail(x)) to.push(String(x).trim()); });
  if(isEmail(ADMIN_EMAIL)) to.push(ADMIN_EMAIL);

  var seen = {}, list = [];
  to.forEach(function(a){ var k = a.toLowerCase(); if(!seen[k]){ seen[k] = 1; list.push(a); } });
  if(!list.length) return {ok:false, error:'no_recipients'};

  var html = minuteHtml(d);
  var name = 'محضر-' + minuteNo;
  var pdf;
  try{
    pdf = Utilities.newBlob(html, 'text/html', name + '.html')
      .getAs('application/pdf').setName(name + '.pdf');
  }catch(err){ pdf = null; }   // البريد يُرسل بالنص المنسّق حتى لو تعذّر توليد PDF

  MailApp.sendEmail({
    to: list.join(','),
    subject: 'محضر اجتماع ' + minuteNo + ' — ' + (f.title || ''),
    htmlBody: html,
    name: SENDER_NAME,
    attachments: pdf ? [pdf] : []
  });
  return {ok:true, sent:list.length, recipients:list, pdf: !!pdf};
}

/* ====================== REMINDERS — DUE DECISIONS ==================== */
function toDate(v){
  if(v instanceof Date) return v;
  var s = String(v || '').trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  var p = s.split('-');
  return new Date(+p[0], +p[1] - 1, +p[2]);
}
function daysBetween(a, b){
  return Math.round((a.getTime() - b.getTime()) / 86400000);
}

/** every open decision, newest due first */
function openDecisions(){
  var sh = decisionsSheet(), last = sh.getLastRow();
  if(last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, DECISION_HEADERS.length).getValues();
  var today = new Date(); today.setHours(0,0,0,0);
  var out = [];
  vals.forEach(function(v, i){
    if(isDone(v[7])) return;
    var due = toDate(v[6]);
    out.push({
      row: i + 2, minuteNo: v[0], seq: v[2], decision: v[3], owner: v[4],
      ownerEmail: v[5], due: v[6], status: v[7], meeting: v[8],
      daysLeft: due ? daysBetween(due, today) : null
    });
  });
  return out;
}

/** run daily — reminds each owner, then sends the admin a digest */
function sendReminders(){
  var items = openDecisions().filter(function(x){
    return x.daysLeft !== null && x.daysLeft <= REMIND_DAYS;
  });
  if(!items.length) return;

  var sh = decisionsSheet();
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var byOwner = {};

  items.forEach(function(x){
    if(!isEmail(x.ownerEmail)) return;
    (byOwner[x.ownerEmail] = byOwner[x.ownerEmail] || []).push(x);
  });

  Object.keys(byOwner).forEach(function(email){
    var rows = byOwner[email];
    MailApp.sendEmail({
      to: email,
      subject: 'تذكير: ' + rows.length + ' مهمة متابعة مستحقة — محاضر اجتماعات ICTD',
      htmlBody: remindHtml(rows[0].owner, rows, false),
      name: SENDER_NAME
    });
    rows.forEach(function(x){ sh.getRange(x.row, 11).setValue(stamp); });
  });

  if(isEmail(ADMIN_EMAIL)){
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: 'ملخص متابعة القرارات — ' + items.length + ' مهمة مستحقة أو متأخرة',
      htmlBody: remindHtml('', items, true),
      name: SENDER_NAME
    });
  }
}

function remindHtml(owner, rows, isDigest){
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

  rows.sort(function(a,b){ return (a.daysLeft||0) - (b.daysLeft||0); }).forEach(function(x){
    var left = x.daysLeft < 0 ? ('<b style="color:#9d1f18">متأخر ' + Math.abs(x.daysLeft) + ' يوم</b>')
             : x.daysLeft === 0 ? '<b style="color:#8a5b06">اليوم</b>'
             : (x.daysLeft + ' يوم');
    h += '<tr><td>' + esc(x.decision) + '</td>'
      + (isDigest ? '<td>' + esc(x.owner) + '</td>' : '')
      + '<td>' + esc(x.minuteNo) + '</td><td>' + esc(x.due) + '</td>'
      + '<td>' + esc(x.status) + '</td><td>' + left + '</td></tr>';
  });

  return h + '</table></div><div class="foot">رسالة آلية من نظام محاضر اجتماعات ICTD — '
    + 'يرجى تحديث حالة المهمة في محضر الاجتماع المعني.</div></div></body></html>';
}

/* ============================== SETUP =============================== */
function setup(){
  minutesSheet();
  decisionsSheet();
}

/** run once to schedule the daily reminder at 07:00 Riyadh time */
function installReminderTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReminders').timeBased().atHour(7).everyDays(1).create();
}

/** optional helper — check what today's reminder run would send */
function previewReminders(){
  var items = openDecisions().filter(function(x){
    return x.daysLeft !== null && x.daysLeft <= REMIND_DAYS;
  });
  Logger.log(JSON.stringify(items, null, 2));
}
