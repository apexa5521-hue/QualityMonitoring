// ============================================================
//  APEXCARE Quality Management System — Google Apps Script
//  Version 2.0 | مدير الجودة: إبراهيم
// ============================================================

const SPREADSHEET_ID = ''; // اتركه فارغاً للعمل مع الشيت الحالي
const SHEETS = {
  ACHIEVEMENTS: 'Achievements',
  AUDIT:        'FieldAudits',
  DAILY:        'DailyTracker',
  MAINTENANCE:  'Maintenance',
  CBAHI:        'CBAHI_Meetings',
  INVENTORY:    'Inventory',
  DASHBOARD:    'Dashboard'
};

// ============================================================
//  WEB APP ENTRY POINT
// ============================================================
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('APEXCARE Quality System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0');
}

// ============================================================
//  ROUTER
// ============================================================
function handleRequest(payload) {
  try {
    switch (payload.action) {
      case 'getAchievements':   return getAchievements();
      case 'addAchievement':    return addAchievement(payload.data);
      case 'updateAchievement': return updateAchievement(payload.rowIndex, payload.data);
      case 'deleteAchievement': return deleteAchievement(payload.rowIndex);
      case 'saveAudit':         return saveAudit(payload.data);
      case 'saveDaily':         return saveDaily(payload.data);
      case 'saveMaintenance':   return saveMaintenance(payload.data);
      case 'saveCBAHI':         return saveCBAHI(payload.data);
      case 'lookupDevice':      return lookupDevice(payload.code);
      case 'getDashboard':      return getDashboardData();
      default: return { success: false, message: 'Unknown action' };
    }
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
//  INITIALIZATION — نفّذ مرة واحدة
// ============================================================
function initializeSystem() {
  const ss = getSpreadsheet();
  ensureSheet(ss, SHEETS.ACHIEVEMENTS, ['التاريخ','القسم','وصف الإنجاز','Description EN','نسبة الإنجاز','الحالة','ملاحظات','آخر تعديل']);
  ensureSheet(ss, SHEETS.AUDIT,        ['Timestamp','Branch','الفرع','IPC مكافحة العدوى','Laser Safety','Emergency الطوارئ','Reception الاستقبال','5S بيئة العمل','Notes AR','Notes EN','Inspector']);
  ensureSheet(ss, SHEETS.DAILY,        ['Timestamp','Date','الفرع','Branch','Completed AR','Completed EN','Pending AR','Pending EN']);
  ensureSheet(ss, SHEETS.MAINTENANCE,  ['Timestamp','Branch','Device Code','Device Name','Issue AR','Issue EN','Priority','Status','Reported By']);
  ensureSheet(ss, SHEETS.CBAHI,        ['Timestamp','Meeting Date','Attendees','Outcomes AR','Outcomes EN','Action Items','Next Meeting']);
  ensureSheet(ss, SHEETS.INVENTORY,    ['Device Code','Device Name AR','Device Name EN','Branch','Department','Status','Last Maintenance']);
  ensureSheet(ss, SHEETS.DASHBOARD,    ['Metric','Value','Last Updated']);
  populateSampleAchievements(ss);
  populateSampleInventory(ss);
  buildDashboard(ss);
  return { success: true, message: 'تم تهيئة النظام بنجاح ✅' };
}

// ============================================================
//  ACHIEVEMENTS CRUD
//  Sheet columns (A-H):
//  A=date  B=section  C=desc_ar  D=desc_en  E=progress  F=status  G=notes  H=lastEdit
// ============================================================

function getAchievements() {
  const ss   = getSpreadsheet();
  const sh   = ss.getSheetByName(SHEETS.ACHIEVEMENTS);
  const last = sh.getLastRow();
  if (last < 2) return { success: true, data: [] };

  const rows = sh.getRange(2, 1, last - 1, 8).getValues();
  const data = rows.map((r, i) => ({
    rowIndex: i + 2,
    date:     r[0] instanceof Date ? Utilities.formatDate(r[0],'Asia/Riyadh','yyyy-MM-dd') : String(r[0]),
    section:  r[1],
    desc_ar:  r[2],
    desc_en:  r[3],
    progress: r[4],
    status:   r[5],
    notes:    r[6],
    lastEdit: r[7] instanceof Date ? Utilities.formatDate(r[7],'Asia/Riyadh','yyyy-MM-dd HH:mm') : String(r[7])
  }));
  return { success: true, data };
}

function addAchievement(d) {
  const ss  = getSpreadsheet();
  const sh  = ss.getSheetByName(SHEETS.ACHIEVEMENTS);
  sh.appendRow([d.date, d.section, d.desc_ar, d.desc_en||'', d.progress||'100%', d.status||'مكتمل', d.notes||'', new Date()]);
  styleAchSheet(sh);
  updateDashboard(ss);
  return { success: true, message: 'تمت إضافة الإنجاز ✅' };
}

function updateAchievement(rowIndex, d) {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.ACHIEVEMENTS);
  sh.getRange(rowIndex, 1, 1, 8).setValues([[
    d.date, d.section, d.desc_ar, d.desc_en||'', d.progress||'100%', d.status||'مكتمل', d.notes||'', new Date()
  ]]);
  return { success: true, message: 'تم تحديث الإنجاز ✅' };
}

function deleteAchievement(rowIndex) {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.ACHIEVEMENTS);
  sh.deleteRow(rowIndex);
  updateDashboard(ss);
  return { success: true, message: 'تم حذف الإنجاز 🗑' };
}

function styleAchSheet(sh) {
  sh.getRange(1,1,1,8).setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1,8); } catch(e){}
}

function populateSampleAchievements(ss) {
  const sh = ss.getSheetByName(SHEETS.ACHIEVEMENTS);
  if (sh.getLastRow() > 1) return;
  const now = new Date();
  const rows = [
    ['2026-02-09','اعتماد سباهي / CBAHI','التواصل مع فريق سباهي وصياغة اتفاقية التعاون وتسويقها للإدارة','Contacted SBAHI team, drafted cooperation agreement','100%','مكتمل','',now],
    ['2026-02-10','اعتماد سباهي / CBAHI','الاجتماع التدشيني مع فريق سباهي — شرح المبالغ والخطط والجدول الزمني','Kickoff meeting — budget, plans and timeline explained','100%','مكتمل','',now],
    ['2026-02-15','تحليل الفجوات / Gap Analysis','زيارات ميدانية لتقييم الفجوات ومراجعة ملاحظات الفريق وتعيين المسؤولين','Field visits for gap analysis, department owners assigned','100%','مكتمل','',now],
    ['2026-02-25','الامتثال الطبي / Compliance','مخاطبة الشركات لتوفير رخص تسويق الأجهزة (MDMA) ورخص الاستيراد لهيئة الغذاء والدواء','Contacted suppliers for MDMA licenses and SFDA import permits','90%','جارٍ','قيد الاكتمال — آخر رد متوقع نهاية مارس',now],
    ['2026-03-05','إدارة الأصول / Assets','جرد كامل لأجهزة فرع بريدة ووضع ملصقات التكويد QR','Full inventory of Buraydah devices with QR asset tags','100%','مكتمل','',now],
    ['2026-03-10','الكول سنتر / Call Center','بناء نظام حوافز للمشرفة رهف وفريقها (9 موظفات)','Incentive system for supervisor Rahaf and 9 agents','100%','مكتمل','نقطتان لكل رصد — حافز 5 ريال — سقف = عدد الموظفات × 100',now],
    ['2026-03-15','الأدلة الرقمية / Digital','دليل إرشادي للموظفين مع أسعار الخدمات محمي بكلمة مرور ونظام الردود الموحدة','Password-protected employee guide and unified response system','100%','مكتمل','',now],
    ['2026-03-20','الموارد البشرية / HR','مراجعة الوصف الوظيفي مع سباهي وتفعيل نماذج Orientation واستبيانات رضا الموظفين','Job descriptions reviewed, orientation forms and surveys activated','100%','مكتمل','',now],
  ];
  sh.getRange(2,1,rows.length,8).setValues(rows);
  styleAchSheet(sh);
}

// ============================================================
//  OTHER SAVES
// ============================================================
function saveAudit(d) {
  const ss=getSpreadsheet(); const sh=ss.getSheetByName(SHEETS.AUDIT);
  sh.appendRow([new Date(),d.branch_en,d.branch_ar,d.sterilization,d.laser,d.emergency,d.reception,d.notes_ar,d.notes_en,'Quality Manager']);
  updateDashboard(ss);
  return {success:true,message:'تم حفظ نتيجة الزيارة ✅'};
}
function saveDaily(d) {
  const ss=getSpreadsheet(); const sh=ss.getSheetByName(SHEETS.DAILY); const now=new Date();
  sh.appendRow([now,Utilities.formatDate(now,'Asia/Riyadh','yyyy-MM-dd'),d.branch_ar,d.branch_en,d.completed_ar,d.completed_en,d.pending_ar,d.pending_en]);
  return {success:true,message:'تم حفظ متابعة اليوم ✅'};
}
function saveMaintenance(d) {
  const ss=getSpreadsheet(); const sh=ss.getSheetByName(SHEETS.MAINTENANCE);
  sh.appendRow([new Date(),d.branch,d.deviceCode,d.deviceName||lookupDeviceName(d.deviceCode),d.issue_ar,d.issue_en,d.priority||'عادي','مفتوح','Quality Manager']);
  return {success:true,message:'تم تسجيل بلاغ الصيانة ✅'};
}
function saveCBAHI(d) {
  const ss=getSpreadsheet(); const sh=ss.getSheetByName(SHEETS.CBAHI);
  sh.appendRow([new Date(),d.meetingDate,d.attendees,d.outcomes_ar,d.outcomes_en,d.actionItems,d.nextMeeting]);
  return {success:true,message:'تم حفظ مخرجات الاجتماع ✅'};
}

// ============================================================
//  DEVICE LOOKUP
// ============================================================
function lookupDevice(code) {
  if(!code) return {found:false};
  const data=getSpreadsheet().getSheetByName(SHEETS.INVENTORY).getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).trim().toUpperCase()===String(code).trim().toUpperCase())
      return{found:true,code:data[i][0],name_ar:data[i][1],name_en:data[i][2],branch:data[i][3],department:data[i][4],status:data[i][5]};
  }
  return {found:false};
}
function lookupDeviceName(code){const r=lookupDevice(code);return r.found?r.name_ar:code;}

// ============================================================
//  DASHBOARD
// ============================================================
function getDashboardData() {
  const ss=getSpreadsheet();
  const achRows=Math.max(0,ss.getSheetByName(SHEETS.ACHIEVEMENTS).getLastRow()-1);
  const auditRows=Math.max(0,ss.getSheetByName(SHEETS.AUDIT).getLastRow()-1);
  const maintRows=Math.max(0,ss.getSheetByName(SHEETS.MAINTENANCE).getLastRow()-1);
  const dailyRows=Math.max(0,ss.getSheetByName(SHEETS.DAILY).getLastRow()-1);
  let completed=0;
  if(achRows>0){ss.getSheetByName(SHEETS.ACHIEVEMENTS).getRange(2,6,achRows,1).getValues().forEach(r=>{if(r[0]==='مكتمل')completed++;});}
  let pass=0,total=0;
  if(auditRows>0){ss.getSheetByName(SHEETS.AUDIT).getRange(2,4,auditRows,5).getValues().forEach(row=>row.forEach(c=>{total++;if(c===true)pass++;}));}
  return{success:true,achCount:achRows,achCompleted:completed,auditCount:auditRows,maintCount:maintRows,dailyCount:dailyRows,auditPassRate:total>0?Math.round(pass/total*100):0};
}
function buildDashboard(ss){
  const sh=ss.getSheetByName(SHEETS.DASHBOARD);sh.clearContents();
  sh.getRange(1,1,1,3).setValues([['Metric','Value','Last Updated']]).setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(2,1,5,3).setValues([['Total Achievements',0,new Date()],['Completed',0,new Date()],['Field Audits',0,new Date()],['Maintenance',0,new Date()],['Audit Pass %',0,new Date()]]);
}
function updateDashboard(ss){
  try{const d=getDashboardData();const sh=ss.getSheetByName(SHEETS.DASHBOARD);const now=new Date();
  sh.getRange(2,2).setValue(d.achCount);sh.getRange(3,2).setValue(d.achCompleted);
  sh.getRange(4,2).setValue(d.auditCount);sh.getRange(5,2).setValue(d.maintCount);
  sh.getRange(6,2).setValue(d.auditPassRate);sh.getRange(2,3,5,1).setValue(now);}catch(e){}
}

// ============================================================
//  INVENTORY SAMPLE
// ============================================================
function populateSampleInventory(ss){
  const sh=ss.getSheetByName(SHEETS.INVENTORY);if(sh.getLastRow()>1)return;
  const inv=[['APX-BRD-XRAY-001','جهاز أشعة سيني','Dental X-Ray','بريدة','أسنان','نشط',''],['APX-BRD-UNIT-001','يونيت الأسنان #1','Dental Unit #1','بريدة','أسنان','نشط',''],['APX-BRD-UNIT-002','يونيت الأسنان #2','Dental Unit #2','بريدة','أسنان','نشط',''],['APX-BRD-LASR-001','جهاز الليزر الجلدي','Dermatology Laser','بريدة','جلدية','نشط',''],['APX-BRD-AUTO-001','أوتوكلاف التعقيم','Autoclave','بريدة','تعقيم','نشط',''],['APX-UNZ-UNIT-001','يونيت عنيزة #1','Dental Unit Unaizah','عنيزة','أسنان','نشط',''],['APX-UNZ-LASR-001','ليزر عنيزة','Laser Unaizah','عنيزة','جلدية','نشط',''],['APX-BRD-CBCT-001','CBCT ثلاثي الأبعاد','3D CBCT','بريدة','أسنان','نشط','']];
  sh.getRange(2,1,inv.length,7).setValues(inv);
  sh.getRange(1,1,1,7).setBackground('#0D47A1').setFontColor('#FFFFFF').setFontWeight('bold');sh.setFrozenRows(1);
}

// ============================================================
//  UTILITIES
// ============================================================
function getSpreadsheet(){return SPREADSHEET_ID?SpreadsheetApp.openById(SPREADSHEET_ID):SpreadsheetApp.getActiveSpreadsheet();}
function ensureSheet(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh){sh=ss.insertSheet(name);sh.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#1565C0').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');sh.setFrozenRows(1);}return sh;}
