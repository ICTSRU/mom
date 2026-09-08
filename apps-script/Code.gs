/**
 * ICTD Meeting Minutes — Google Apps Script backend
 * الإدارة التنفيذية للاتصالات وتقنية المعلومات — جامعة سليمان الراجحي
 *
 * Deploy: Extensions ▸ Apps Script ▸ paste this file ▸ Deploy ▸ New deployment
 *         Type: Web app · Execute as: Me · Who has access: Anyone
 * Then copy the /exec URL into the form (زر «إعدادات الحفظ»).
 *
 * Endpoints
 *   GET  ?action=next&year=2026   -> { ok, minuteNo }   next free minute number
 *   GET  ?action=list             -> { ok, rows }       minute index (no payloads)
 *   GET  ?action=get&no=ICTD-...  -> { ok, data }       one full minute (JSON payload)
 *   POST { ...form JSON }         -> { ok, minuteNo }   upsert minute + rebuild its decisions
 */

var SHEET_ID       = '1pw9XIBc7Zt6uF3wJO_gOZQ_meAM9pCqKMMDWsSHEjZ4';
var MINUTES_SHEET  = 'Minutes';
var DECISIONS_SHEET= 'Decisions';
var PREFIX         = 'ICTD-MOM';

var MINUTE_HEADERS = ['Timestamp','MinuteNo','Title','Type','Place','DateG','DateH','Day','From','To',
  'Host','Chair','Secretary','Classification','Attendees','Apologies','PreviousFollowUp','Agenda',
  'Decisions','Notes','NextMeeting','PayloadJSON'];
var DECISION_HEADERS = ['MinuteNo','DateG','Seq','Decision','Owner','DueDate','Status','MeetingTitle','UpdatedAt'];

/* ------------------------------------------------------------------ utils */
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ss(){ return SpreadsheetApp.openById(SHEET_ID); }

function sheetFor(name, headers){
  var book = ss();
  var sh = book.getSheetByName(name);
  if(!sh){
    // the CSV import names the first tab after the file — reuse it for Minutes
    if(name === MINUTES_SHEET && book.getSheets().length === 1 && book.getSheets()[0].getLastRow() <= 1){
      sh = book.getSheets()[0];
      sh.setName(name);
    } else {
      sh = book.insertSheet(name);
    }
  }
  if(sh.getLastRow() === 0 || String(sh.getRange(1,1).getValue()).trim() === ''){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length)
    .setFontWeight('bold').setBackground('#f0ebf8').setFontColor('#3a1464');
  return sh;
}
function minutesSheet(){   return sheetFor(MINUTES_SHEET,   MINUTE_HEADERS); }
function decisionsSheet(){ return sheetFor(DECISIONS_SHEET, DECISION_HEADERS); }

function pad3(n){ n = String(n); while(n.length < 3) n = '0' + n; return n; }
function makeNo(year, seq){ return PREFIX + '-' + year + '-' + pad3(seq); }

/** highest sequence already recorded for a year */
function maxSeq(year){
  var sh = minutesSheet();
  var last = sh.getLastRow();
  if(last < 2) return 0;
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  var re = new RegExp('^' + PREFIX + '-' + year + '-(\\d+)$');
  var max = 0;
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

/** row index (1-based) of a minute number, or 0 */
function findRow(minuteNo){
  var sh = minutesSheet();
  var last = sh.getLastRow();
  if(last < 2) return 0;
  var col = sh.getRange(2, 2, last - 1, 1).getValues();
  for(var i = 0; i < col.length; i++){
    if(String(col[i][0]).trim() === String(minuteNo).trim()) return i + 2;
  }
  return 0;
}

/** flatten a table (array of row-arrays) into readable multi-line text */
function flatten(rows){
  if(!rows || !rows.length) return '';
  return rows
    .filter(function(r){ return r.join('').trim() !== ''; })
    .map(function(r, i){ return (i + 1) + '. ' + r.filter(String).join(' | '); })
    .join('\n');
}

/* ------------------------------------------------------------------- GET */
function doGet(e){
  var p = (e && e.parameter) || {};
  var action = String(p.action || '').toLowerCase();
  try{
    if(action === 'next') return json({ok:true, minuteNo: nextNo(p.year)});
    if(action === 'get'){
      var r = findRow(p.no);
      if(!r) return json({ok:false, error:'not_found'});
      var payload = minutesSheet().getRange(r, MINUTE_HEADERS.length, 1, 1).getValue();
      return json({ok:true, data: payload ? JSON.parse(payload) : null});
    }
    if(action === 'list'){
      var sh = minutesSheet(), last = sh.getLastRow();
      if(last < 2) return json({ok:true, rows:[]});
      var vals = sh.getRange(2, 1, last - 1, 14).getValues();
      return json({ok:true, rows: vals.map(function(v){
        return {timestamp:v[0], minuteNo:v[1], title:v[2], type:v[3], dateG:v[5], chair:v[11]};
      })});
    }
    return json({ok:true, service:'ICTD Meeting Minutes API', version:'1.5'});
  }catch(err){ return json({ok:false, error:String(err)}); }
}

/* ------------------------------------------------------------------ POST */
function doPost(e){
  var lock = LockService.getScriptLock();
  try{
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    var f = data.fields || {};
    var t = data.tables || {};

    var year = String(f.dateG || '').slice(0,4);
    if(!/^\d{4}$/.test(year)) year = String(new Date().getFullYear());

    var minuteNo = String(f.docNo || '').trim();
    if(!minuteNo) minuteNo = nextNo(year);

    var row = [
      new Date(), minuteNo, f.title || '', f.type || '', f.place || '',
      f.dateG || '', f.dateH || '', f.day || '', f.from || '', f.to || '',
      f.host || '', f.chair || '', f.secretary || '', f.classification || '',
      flatten(t.attendees), flatten(t.apologies), flatten(t.previous),
      flatten(t.agenda), flatten(t.decisions), f.notes || '',
      [f.nextDate || '', f.nextTime || '', f.nextPlace || ''].filter(String).join(' — '),
      JSON.stringify(data)
    ];

    var sh = minutesSheet();
    var r = findRow(minuteNo);
    var created = false;
    if(r){ sh.getRange(r, 1, 1, row.length).setValues([row]); }
    else { sh.appendRow(row); created = true; }

    syncDecisions(minuteNo, f, t.decisions || []);

    return json({ok:true, minuteNo: minuteNo, created: created});
  }catch(err){
    return json({ok:false, error:String(err)});
  }finally{
    try{ lock.releaseLock(); }catch(x){}
  }
}

/** rebuild the decision rows belonging to one minute */
function syncDecisions(minuteNo, f, rows){
  var sh = decisionsSheet();
  var last = sh.getLastRow();
  if(last >= 2){
    var col = sh.getRange(2, 1, last - 1, 1).getValues();
    for(var i = col.length - 1; i >= 0; i--){
      if(String(col[i][0]).trim() === minuteNo) sh.deleteRow(i + 2);
    }
  }
  var out = [];
  (rows || []).forEach(function(r, i){
    if(!r || String(r[0] || '').trim() === '') return;
    out.push([minuteNo, f.dateG || '', i + 1, r[0] || '', r[1] || '', r[2] || '',
              r[3] || '', f.title || '', new Date()]);
  });
  if(out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, DECISION_HEADERS.length).setValues(out);
}

/** run once from the editor to create both tabs and their headers */
function setup(){
  minutesSheet();
  decisionsSheet();
}
