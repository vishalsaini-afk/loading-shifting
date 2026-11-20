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
   • shared billing details (company, GSTIN, PAN, bank…) — same on every phone
   • a shared invoice register ("Labour Invoices" tab) so invoice numbers
     never clash, whoever raises the invoice
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
    return llOut_({ ok: true, config: JSON.parse(props.getProperty('LL_CONFIG') || '{}'), invoices: llReadInv_() });
  }

  if (b.action === 'setConfig') {
    var keep = ['co','prefix','gstin','pan','aadhaar','acName','bankName','acNo','ifsc'], c = {};
    keep.forEach(function (k) { c[k] = String((b.config || {})[k] || ''); });
    props.setProperty('LL_CONFIG', JSON.stringify(c));
    return llOut_({ ok: true, config: c });
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
