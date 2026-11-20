/* =====================================================================
   LABOUR LOG — SHARED ACCESS ADD-ON
   Paste this whole file into your EXISTING Labour Log Apps Script project
   (File ▸ + ▸ Script, name it "shared"). Do NOT replace your old code.

   Then add ONE line as the very first line inside your existing doPost(e):

       var s = llShared_(e); if (s) return s;

   Deploy ▸ Manage deployments ▸ Edit ▸ Version: New version ▸ Deploy
   (the /exec URL stays the same).

   What it adds:
   • one common Access PIN for everyone (set LL_ACCESS_PIN below)
   • shared contractor list in a "Labour Contractors" tab (name, GSTIN, PAN,
     Aadhaar, bank…) — entered once, picked on the invoice; edit it in the
     sheet or the app, both stay in step
   • a shared invoice register ("Labour Invoices" tab) so invoice numbers
     never clash; delete a row there and that number is free again
   Your existing entry sheet and columns are untouched.
   ===================================================================== */

var LL_ACCESS_PIN = '';   // e.g. '4321'. Leave '' for no PIN.
var LL_SS_ID      = '';   // leave '' if this script is attached to the sheet; else paste the Sheet ID
var LL_INV_TAB    = 'Labour Invoices';
var LL_INV_HEAD   = ['Logged At','Invoice No','Invoice Date','Period','Purpose','Reason','Filter',
                     'Billed By','Billed To','Entries','Subtotal','GST %','Amount','Raised By'];

function llShared_(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return null; }
  if (!b || !b.action) return null;

  // one common PIN — checked on every call (list, save, delete, translate…)
  if (LL_ACCESS_PIN && String(b.pin || '') !== String(LL_ACCESS_PIN)) {
    return llOut_({ ok: false, error: 'pin' });
  }

  var props = PropertiesService.getScriptProperties();

  if (b.action === 'config') {
    var cfg = JSON.parse(props.getProperty('LL_CONFIG') || '{}');
    delete cfg.contractors;
    var cs = llReadContr_();
    // first run: move a list saved earlier in script storage into the sheet tab
    var old = JSON.parse(props.getProperty('LL_CONFIG') || '{}').contractors;
    if (!cs.length && Array.isArray(old) && old.length) { llWriteContr_(old); cs = llReadContr_(); }
    cfg.contractors = cs;
    return llOut_({ ok: true, config: cfg, invoices: llReadInv_() });
  }

  if (b.action === 'setConfig') {
    var src = b.config || {}, c = { co: String(src.co || ''), prefix: String(src.prefix || '') };
    // contractor list — entered once, picked on every invoice
    var F = ['id','name','addr','gstin','pan','aadhaar','acName','bankName','acNo','ifsc','prefix'];
    c.contractors = (Array.isArray(src.contractors) ? src.contractors : []).map(function (x) {
      var o = {}; F.forEach(function (k) { o[k] = String((x || {})[k] || ''); }); return o;
    }).filter(function (o) { return o.id && o.name; });
    var lk = LockService.getScriptLock(); lk.waitLock(20000);
    try { llWriteContr_(c.contractors); } finally { lk.releaseLock(); }
    props.setProperty('LL_CONFIG', JSON.stringify({ co: c.co, prefix: c.prefix }));
    return llOut_({ ok: true });
  }

  if (b.action === 'logInvoice') {
    var i = b.invoice || {};
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var sh = llInvSheet_();
      var dup = llReadInv_().some(function (r) { return r.no === String(i.no); });
      if (!dup) {
        sh.appendRow([new Date(), String(i.no || ''), i.date || '', i.period || '', i.purpose || 'Both',
          i.reason || 'Both', i.filter || '', i.billedBy || '', i.billedTo || '', Number(i.entries) || 0,
          Number(i.subtotal) || 0, Number(i.gst) || 0, Number(i.amount) || 0, i.by || b.user || '']);
      }
      return llOut_({ ok: true, duplicate: dup });
    } finally { lock.releaseLock(); }
  }

  return null;   // anything else → your existing doPost code handles it
}

/* ---- contractors live in a sheet tab, so the office can edit them there ---- */
var LL_CONTR_TAB  = 'Labour Contractors';
var LL_CONTR_HEAD = ['ID','Name','Address & Phone','GSTIN','PAN','Aadhaar','Account Holder','Bank','Account No','IFSC','Invoice Prefix'];
var LL_CONTR_KEYS = ['id','name','addr','gstin','pan','aadhaar','acName','bankName','acNo','ifsc','prefix'];

function llContrSheet_() {
  var ss = llSS_(), sh = ss.getSheetByName(LL_CONTR_TAB);
  if (!sh) {
    sh = ss.insertSheet(LL_CONTR_TAB); sh.appendRow(LL_CONTR_HEAD); sh.setFrozenRows(1);
    sh.getRange('A:K').setNumberFormat('@');   // keep A/C no, Aadhaar etc. as text (no 1.23E+11)
  }
  return sh;
}
function llReadContr_() {
  var sh = llContrSheet_(), v = sh.getDataRange().getDisplayValues(), out = [], fix = false;
  for (var r = 1; r < v.length; r++) {
    if (!String(v[r][1]).trim()) continue;             // skip rows without a name
    if (!String(v[r][0]).trim()) {                    // row typed in the sheet without ID → give it one
      v[r][0] = 'c' + Date.now().toString(36) + r; sh.getRange(r + 1, 1).setValue(v[r][0]); fix = true;
    }
    var o = {}; LL_CONTR_KEYS.forEach(function (k, i) { o[k] = String(v[r][i] || '').trim(); });
    out.push(o);
  }
  return out;
}
function llWriteContr_(list) {
  var sh = llContrSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, LL_CONTR_HEAD.length).clearContent();
  var rows = (list || []).map(function (x) { return LL_CONTR_KEYS.map(function (k) { return String((x || {})[k] || ''); }); });
  if (rows.length) sh.getRange(2, 1, rows.length, LL_CONTR_HEAD.length).setNumberFormat('@').setValues(rows);
}

function llSS_() { return LL_SS_ID ? SpreadsheetApp.openById(LL_SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function llInvSheet_() {
  var ss = llSS_(), sh = ss.getSheetByName(LL_INV_TAB);
  if (!sh) { sh = ss.insertSheet(LL_INV_TAB); sh.appendRow(LL_INV_HEAD); sh.setFrozenRows(1); }
  return sh;
}

function llReadInv_() {
  var v = llInvSheet_().getDataRange().getValues(); v.shift();
  return v.filter(function (r) { return r[1]; }).map(function (r) {
    var d = r[2] instanceof Date ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[2]);
    return { no: String(r[1]), date: d, period: String(r[3]), amount: Number(r[12]) || 0, by: String(r[13] || '') };
  });
}

function llOut_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
