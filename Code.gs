// CSI Monaghan Church – Family Form backend (Google Apps Script)
// Paste into script.google.com (Extensions > Apps Script from your Google Sheet).
const ADMIN_KEY = 'change-this-password'; // <-- set your own admin password

const FAMILY_COLS = ['FamilyID','Submitted','HeadName','FatherHusbandName','DOB','Age','BaptismDate','MarriageDate',
  'Qualification','Occupation','Mobile','Email','BloodGroup','Address','Area','Pincode',
  'WifeName','WifeDOB','WifeAge','WifeBaptism','WifeQualification','WifeOccupation','WifeMobile',
  'NumChildren','Remarks'];
const CHILD_COLS = ['FamilyID','HeadName','ChildName','Gender','DOB','Age','BaptismDate','Qualification','Mobile','Address'];

function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); sh.getRange(1,1,1,cols.length).setFontWeight('bold'); }
  return sh;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const fam = sheet_('Families', FAMILY_COLS), ch = sheet_('Children', CHILD_COLS);
    const id = 'CSI-' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyMMdd') + '-' + String(fam.getLastRow()).padStart(4, '0');
    const f = d.family || {}; f.FamilyID = id; f.Submitted = new Date(); f.NumChildren = (d.children || []).length;
    fam.appendRow(FAMILY_COLS.map(c => f[c] || ''));
    (d.children || []).forEach(k => { k.FamilyID = id; k.HeadName = f.HeadName; ch.appendRow(CHILD_COLS.map(c => k[c] || '')); });
    return json_({ ok: true, id: id });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}

function doGet(e) {
  if (!e.parameter || e.parameter.key !== ADMIN_KEY) return json_({ ok: false, error: 'Unauthorized' });
  const read = (name, cols) => { const sh = sheet_(name, cols), v = sh.getDataRange().getValues(); const h = v.shift();
    return v.map(r => Object.fromEntries(h.map((c, i) => [c, r[i] instanceof Date ? Utilities.formatDate(r[i], 'Asia/Kolkata', 'yyyy-MM-dd HH:mm') : r[i]]))); };
  return json_({ ok: true, families: read('Families', FAMILY_COLS), children: read('Children', CHILD_COLS) });
}
