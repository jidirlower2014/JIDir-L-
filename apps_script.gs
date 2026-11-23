/**
 * جماعت اسلامی ضلع دیر پائین — Google Sheet بیک اینڈ
 * (روپے جمع کریں + اطلاعات)
 *
 * پہلی بار لگانا:
 * 1. Google Sheet کھولیں  ->  Extensions  ->  Apps Script
 * 2. پرانا کوڈ مٹا کر یہ پورا کوڈ پیسٹ کریں، Save کریں
 * 3. Deploy -> New deployment -> Type: Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. جو "Web app URL" (آخر میں /exec) ملے، وہ index.html میں SHEET_API_URL میں ڈالیں
 *
 * پہلے سے لگا ہوا ہو تو (اپ ڈیٹ کے لیے):
 *    کوڈ بدلیں، Save کریں، پھر Deploy -> Manage deployments -> ✏️ (Edit)
 *    -> Version: "New version" -> Deploy.   لنک وہی رہے گا۔
 *
 * شیٹس خود بنیں گی:
 *   Deposits      : ہر جمع شدہ رقم کی لائن
 *   Hisab         : ہر مقام کی تازہ وصولی/باقی اور 12 مہینے
 *   Notifications : بھیجی گئی اطلاعات (زون، متن، تاریخ، ربن کی آخری تاریخ، PDF)
 *   PDF فائلیں Google Drive کے فولڈر "JID Notification PDFs" میں محفوظ ہوں گی
 *   (نیا کوڈ لگانے کے بعد Deploy پر Google Drive کی اجازت دوبارہ مانگے گا — Allow کر دیں)
 */
// اگر Sheet سے Extensions -> Apps Script نہ کھلے تو script.google.com پر نیا پروجیکٹ بنا کر
// یہاں اپنی Sheet کا ID ڈالیں (Sheet کے لنک میں /d/ اور /edit کے درمیان والا حصہ)
// Sheet کا ID یہاں لکھنے کی ضرورت نہیں: Project Settings -> Script properties میں SHEET_ID کے نام سے ایک بار رکھ دیں،
// پھر نیا کوڈ paste کرنے پر بھی یہ محفوظ رہے گا۔ (Sheet کے اندر سے Extensions -> Apps Script کھولا ہو تو کچھ بھی ڈالنا نہیں)
const SHEET_ID = PropertiesService.getScriptProperties().getProperty('SHEET_ID') || '';

const SECRET = '';   // اختیاری: کوئی خفیہ لفظ رکھیں تو HTML میں SHEET_API_KEY میں بھی وہی رکھیں

const DEP_HEAD = ['id','saved_at','date','mkey','zone','muqam','month','amount'];
const HISAB_HEAD = ['mkey','zone','muqam','ameer','sabiqa','saalana','wasooli','baqaya',
  'jul','aug','sep','oct','nov','dec','jan','feb','mar','apr','may','jun'];
const NOTIF_HEAD = ['id','created','zone','text','tickerUntil','fileId','fileName'];
const CAND_HEAD = ['id','zone','name','phone','muqam','status','note','created','t','father','vc','att'];
const ARK_HEAD = ['id','type','reason','zone','zoneName','muqam','name','father','phone','key','date','t','extra'];
const FIN_HEAD = ['id','mkey','zone','zoneName','muqam','ameer','contact','sabiqa','saalana','wasooli','date','t','ex','sm'];
const RF_HEAD = ['id','kind','title','from','to','fields','created','t'];
const RD_HEAD = ['id','form','zone','data','st','t'];
const PDF_FOLDER = 'JID Notification PDFs';

function sheet_(name, head){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); }
  return sh;
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
// نوٹیفکیشن شیٹ کے پرانے ورژن میں fileId/fileName کے کالم نہ ہوں تو بنا دیں
function notifSheet_(){
  const sh = sheet_('Notifications', NOTIF_HEAD);
  if(sh.getRange(1,6).getValue() !== 'fileId'){ sh.getRange(1,6,1,2).setValues([['fileId','fileName']]); }
  return sh;
}
function pdfFolder_(){
  const it = DriveApp.getFoldersByName(PDF_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PDF_FOLDER);
}
function ids_(sh){
  return sh.getRange(1,1,Math.max(sh.getLastRow(),1),1).getValues().map(r=>String(r[0]));
}

function doGet(e){
  try{ return doGet_(e); }catch(err){ return json_({error:String(err && err.message || err)}); }
}
function doGet_(e){
  if(SECRET && (!e.parameter || e.parameter.key !== SECRET)) return json_({error:'forbidden'});

  const dep = sheet_('Deposits', DEP_HEAD);
  const deposits = dep.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]), date:String(r[2]), mkey:String(r[3]), month:String(r[6]), amount:Number(r[7])
  }));

  const ns = notifSheet_();
  const notifications = ns.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]), created:Number(r[1]), zone:String(r[2]), text:String(r[3]), tickerUntil:Number(r[4])||0,
    fileId:String(r[5]||''), fileName:String(r[6]||'')
  }));

  const cs = sheet_('Candidates', CAND_HEAD);
  const candidates = cs.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]), zone:String(r[1]), name:String(r[2]), phone:String(r[3]), muqam:String(r[4]),
    status:String(r[5]), note:String(r[6]), created:Number(r[7]), t:Number(r[8]), father:String(r[9]||''), vc:String(r[10]||''), att:String(r[11]||'')
  }));

  const as = sheet_('ArkanLog', ARK_HEAD);
  const arkLog = as.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]), type:String(r[1]), reason:String(r[2]), zone:String(r[3]), muqam:String(r[5]), n:String(r[6]),
    f:String(r[7]), c:String(r[8]), key:String(r[9]), date:String(r[10]), t:Number(r[11]), x:String(r[12]||'')
  }));

  const fs = sheet_('FinEdits', FIN_HEAD);
  const finEdits = fs.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]), mkey:String(r[1]), ameer:String(r[5]), contact:String(r[6]),
    sabiqa:Number(r[7]), saalana:Number(r[8]), wasooli:Number(r[9]), t:Number(r[11]), ex:Number(r[12])||0, sm:(r[13]===undefined||r[13]==='') ? '' : Number(r[13])
  }));

  const rfs = sheet_('ReportForms', RF_HEAD), rds = sheet_('ReportData', RD_HEAD);
  const reportForms = rfs.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]), kind:String(r[1]), title:String(r[2]), from:String(r[3]), to:String(r[4]), fields:String(r[5]||'[]'), created:Number(r[6]), t:Number(r[7])}));
  const reportData = rds.getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({id:String(r[0]), form:String(r[1]), zone:String(r[2]), data:String(r[3]||'{}'), st:String(r[4]), t:Number(r[5])}));

  return json_({deposits, notifications, candidates, arkLog, finEdits, reportForms, reportData});
}

function doPost(e){
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    const body = JSON.parse(e.postData.contents);
    if(SECRET && body.key !== SECRET) return json_({error:'forbidden'});

    // ---------- اطلاع بھیجنا ----------
    if(body.action === 'notif_add'){
      const n = body.n;
      if(!n || !n.id || !n.text) return json_({error:'bad request'});
      const ns = notifSheet_();
      if(ids_(ns).indexOf(String(n.id)) === -1){
        let fileId = '';
        if(n.pdf){
          const blob = Utilities.newBlob(Utilities.base64Decode(n.pdf), 'application/pdf', n.fileName || (n.id + '.pdf'));
          const file = pdfFolder_().createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          fileId = file.getId();
        }
        ns.appendRow([n.id, Number(n.created), n.zone, n.text, Number(n.tickerUntil)||0, fileId, n.fileName || '']);
      }
      return json_({ok:true});
    }

    // ---------- اراکین/زونز: ایک ساتھ بہت سے اندراجات (امپورٹ / بحالی) ----------
    if(body.action === 'ark_bulk'){
      const es = body.es || []; if(!es.length) return json_({ok:true});
      const as = sheet_('ArkanLog', ARK_HEAD), have = {}; ids_(as).forEach(x=>{ have[x] = 1; });
      const rows = es.filter(e=>e && e.id && !have[e.id]).map(e=>[e.id, e.type, e.reason||'', e.zone||'', e.zn||'', e.muqam||'', e.n||'', e.f||'', e.c||'', e.key||'', e.date||'', Number(e.t)||0, JSON.stringify(e.x||{})]);
      if(rows.length){
        const start = as.getLastRow() + 1;
        as.getRange(start, 9, rows.length, 1).setNumberFormat('@'); as.getRange(start, 11, rows.length, 1).setNumberFormat('@');
        as.getRange(start, 1, rows.length, ARK_HEAD.length).setValues(rows);
      }
      return json_({ok:true, added:rows.length});
    }

    // ---------- اراکین: اضافہ/حذف کا ریکارڈ (الگ شیٹ ArkanLog) ----------
    if(body.action === 'ark_log'){
      const e = body.e; if(!e || !e.id) return json_({error:'bad request'});
      const as = sheet_('ArkanLog', ARK_HEAD);
      if(ids_(as).indexOf(String(e.id)) === -1){
        as.appendRow([e.id, e.type, e.reason, e.zone, e.zn||'', e.muqam, e.n, e.f||'', "'" + (e.c||''), e.key, "'" + e.date, Number(e.t), JSON.stringify(e.x||{})]);
      }
      return json_({ok:true});
    }

    // ---------- مالیاتی ریکارڈ میں ترمیم (ہر ترمیم الگ لائن، تاریخچہ محفوظ) ----------
    if(body.action === 'fin_set'){
      const e = body.e; if(!e || !e.id || !e.mkey) return json_({error:'bad request'});
      const fs = sheet_('FinEdits', FIN_HEAD);
      if(ids_(fs).indexOf(String(e.id)) === -1){
        fs.appendRow([e.id, e.mkey, e.zone, e.zoneName||'', e.muqam||'', e.ameer||'', "'" + (e.contact||''), Number(e.sabiqa), Number(e.saalana), Number(e.wasooli), "'" + e.date, Number(e.t), Number(e.ex)||0, (e.sm===''||e.sm==null) ? '' : Number(e.sm)]);
      }
      return json_({ok:true});
    }

    // ---------- رپورٹس: فارم / زون کی رپورٹ ----------
    if(body.action === 'rf_save'){
      const f = body.f; if(!f || !f.id) return json_({error:'bad request'});
      const sh = sheet_('ReportForms', RF_HEAD), row = [f.id, f.kind, f.title, "'" + (f.from||''), "'" + (f.to||''), JSON.stringify(f.fields||[]), Number(f.created), Number(f.t)];
      const at = ids_(sh).indexOf(String(f.id));
      if(at === -1) sh.appendRow(row); else sh.getRange(at+1,1,1,row.length).setValues([row]);
      return json_({ok:true});
    }
    if(body.action === 'rf_del'){
      const sh = sheet_('ReportForms', RF_HEAD), at = ids_(sh).indexOf(String(body.id));
      if(at > 0) sh.deleteRow(at + 1);
      return json_({ok:true});
    }
    if(body.action === 'rd_save'){
      const d = body.d; if(!d || !d.id) return json_({error:'bad request'});
      const sh = sheet_('ReportData', RD_HEAD), row = [d.id, d.form, d.zone, JSON.stringify(d.data||{}), d.st, Number(d.t)];
      const at = ids_(sh).indexOf(String(d.id));
      if(at === -1) sh.appendRow(row); else sh.getRange(at+1,1,1,row.length).setValues([row]);
      return json_({ok:true});
    }

    // ---------- امیدوار رکنیت: محفوظ/اپ ڈیٹ ----------
    if(body.action === 'cand_save'){
      const c = body.c; if(!c || !c.id || !c.name) return json_({error:'bad request'});
      const cs = sheet_('Candidates', CAND_HEAD);
      const row = [c.id, c.zone, c.name, "'" + (c.phone||''), c.muqam||'', c.status||'pending', c.note||'', Number(c.created), Number(c.t), c.father||'', c.vc||'', JSON.stringify(c.att||{})];
      const at = ids_(cs).indexOf(String(c.id));
      if(at === -1) cs.appendRow(row); else cs.getRange(at+1,1,1,row.length).setValues([row]);
      return json_({ok:true});
    }
    if(body.action === 'cand_del'){
      const cs = sheet_('Candidates', CAND_HEAD);
      const at = ids_(cs).indexOf(String(body.id));
      if(at > 0) cs.deleteRow(at + 1);
      return json_({ok:true});
    }

    // ---------- اطلاع حذف کرنا ----------
    if(body.action === 'notif_del'){
      const ns = notifSheet_();
      const at = ids_(ns).indexOf(String(body.id));
      if(at > 0){
        const fid = ns.getRange(at + 1, 6).getValue();
        if(fid){ try{ DriveApp.getFileById(String(fid)).setTrashed(true); }catch(e){} }
        ns.deleteRow(at + 1);
      }
      return json_({ok:true});
    }

    // ---------- روپے جمع کریں ----------
    const d = body.deposit, s = body.snapshot;
    if(!d || !d.id || !d.amount) return json_({error:'bad request'});

    const dep = sheet_('Deposits', DEP_HEAD);
    if(ids_(dep).indexOf(String(d.id)) === -1){
      dep.appendRow([d.id, new Date(), "'" + d.date, d.mkey, s.zone, s.name, d.month, d.amount]);
    }

    const hs = sheet_('Hisab', HISAB_HEAD);
    const keys = ids_(hs);
    const mo = s.monthly || {};
    const row = [s.mkey, s.zone, s.name, s.ameer, s.sabiqa, s.saalana, s.wasooli, s.baqaya]
      .concat(['jul','aug','sep','oct','nov','dec','jan','feb','mar','apr','may','jun'].map(k=>mo[k]||0));
    const at = keys.indexOf(String(s.mkey));
    if(at === -1) hs.appendRow(row);
    else hs.getRange(at+1,1,1,row.length).setValues([row]);

    return json_({ok:true});
  } finally { lock.releaseLock(); }
}


/**
 * ایک بار چلائیں (اوپر Run دبائیں، function: setup):
 *  - سب اجازتیں (Sheets + Drive) مانگ لیتا ہے
 *  - سب ٹیبز بنا دیتا ہے
 *  - Execution log میں آپ کی Sheet کا لنک لکھ دیتا ہے
 * نیا کوڈ paste کرنے کے بعد بھی ایک بار یہ چلائیں، پھر Deploy -> New version کریں۔
 */
function setup(){
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error('Sheet نہیں ملی: Project Settings -> Script properties میں SHEET_ID رکھیں');
  sheet_('Deposits', DEP_HEAD); sheet_('Hisab', HISAB_HEAD); notifSheet_();
  sheet_('Candidates', CAND_HEAD); sheet_('ArkanLog', ARK_HEAD); sheet_('FinEdits', FIN_HEAD);
  pdfFolder_();
  Logger.log('ٹھیک ہے ✅ آپ کی Sheet کا لنک: ' + ss.getUrl());
}
