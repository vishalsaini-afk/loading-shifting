/* =====================================================================
   LABOUR LOG — SHARED ADD-ON  (requests · verification · bill numbers)
   Paste this whole file into your EXISTING Labour Log Apps Script project
   (File ▸ + ▸ Script, name it "shared"). Do NOT delete your old code.

   Then add ONE line as the very first line inside your existing doPost(e):

       var s = llShared_(e); if (s) return s;

   Deploy ▸ Manage deployments ▸ Edit ▸ Version: New version ▸ Deploy
   (the /exec URL stays the same).

   What this file handles (your old code still handles translate):
   • one common Access PIN for the Labour Log app (set LL_ACCESS_PIN)
   • the "Labour Log" tab: read / save / delete by COLUMN HEADING, so offer
     numbers never get mixed up. It adds these columns at the end if
     missing: Pulleys, Pulley Days, Entered By, Request ID, Verified By,
     Verified At (columns already there are not moved)
   • "Labour Requests" tab: supervisor requests → approve (rate) / reject
     (reason) → resend
   • supervisor verification of Work / Rework for every new entry
   • "Labour Contractors" tab: contractors + their LAST BILL NO
   • "Labour Invoices" tab: running bill numbers per contractor
   ===================================================================== */

var LL_ACCESS_PIN = '';   // e.g. '4321'. Leave '' for no PIN.
var LL_SS_ID      = '';   // leave '' if this script is attached to the sheet; else paste the Sheet ID
var LL_PULLEY_RATE = 1500;

var LL_LOG_TAB   = 'Labour Log';
var LL_REQ_TAB   = 'Labour Requests';
var LL_INV_TAB   = 'Labour Invoices';
var LL_CONTR_TAB = 'Labour Contractors';

// supervisor actions from Field Tools (hub login, no Labour Log PIN)
var LL_OPEN = { reqCreate: 1, reqMine: 1, reqResend: 1, verifyList: 1, verify: 1 };
// Labour Log app actions handled here (PIN protected)
var LL_MINE = { list: 1, save: 1, 'delete': 1, config: 1, setConfig: 1, logInvoice: 1, reserveInvoice: 1,
                reqList: 1, reqApprove: 1, reqReject: 1 };

function llShared_(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return null; }
  if (!b || !b.action) return null;
  var a = b.action;
  try {
    if (LL_OPEN[a]) return llOut_(llOpen_(a, b));
    if (LL_ACCESS_PIN && String(b.pin || '') !== String(LL_ACCESS_PIN)) return llOut_({ ok: false, error: 'pin' });
    if (LL_MINE[a]) return llOut_(llMine_(a, b));
    return null;   // ping, translate … → your existing doPost code
  } catch (err) {
    return llOut_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ============================ ROUTES ============================ */

function llOpen_(a, b) {
  var sup = String(b.sup || b.user || '').trim();
  if (!sup) return { ok: false, error: 'Supervisor name missing — log in to Field Tools' };

  if (a === 'reqCreate') {
    var r = llCleanReq_(b.req || {}); if (r.error) return { ok: false, error: r.error };
    return llLocked_(function () {
      var id = 'q' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
      r.id = id; r.raisedAt = llNow_(); r.sup = sup; r.status = 'Pending'; r.resends = 0;
      r.history = r.raisedAt + ' raised by ' + sup;
      llReqSheet_().appendRow(llReqRow_(r));
      return { ok: true, id: id, request: r };
    });
  }

  if (a === 'reqMine') {
    var me = sup.toLowerCase();
    var mine = llReadReqs_().filter(function (q) { return q.sup.toLowerCase() === me; });
    mine.sort(function (x, y) { return x.raisedAt < y.raisedAt ? 1 : -1; });
    return { ok: true, requests: mine.slice(0, 200) };
  }

  if (a === 'reqResend') {
    var upd = llCleanReq_(b.req || {}); if (upd.error) return { ok: false, error: upd.error };
    return llLocked_(function () {
      var q = llFindReq_(b.id); if (!q) return { ok: false, error: 'Request not found' };
      if (q.status === 'Approved') return { ok: false, error: 'Already approved' };
      for (var k in upd) q[k] = upd[k];
      q.status = 'Pending'; q.rejectReason = ''; q.decidedBy = ''; q.decidedAt = '';
      q.resends = (Number(q.resends) || 0) + 1;
      q.history = (q.history ? q.history + ' | ' : '') + llNow_() + ' resent by ' + sup;
      llSaveReq_(q);
      return { ok: true, request: q };
    });
  }

  if (a === 'verifyList') {
    var reqs = {}; llReadReqs_().forEach(function (q) { reqs[q.id] = q; });
    var list = llReadLog_().filter(function (x) { return !x.verified; }).map(function (x) {
      x.requestedBy = x.requestId && reqs[x.requestId] ? reqs[x.requestId].sup : '';
      x.mine = x.requestedBy.toLowerCase() === sup.toLowerCase();
      return x;
    });
    return { ok: true, entries: list };
  }

  if (a === 'verify') {
    var reason = String(b.reason || '');
    return llLocked_(function () {
      var m = llLogMap_(true), row = llFindRow_(m, b.id);
      if (!row) return { ok: false, error: 'Entry not found' };
      if (b.problem) {
        var nc = m.C.note, old = String(m.sh.getRange(row, nc + 1).getDisplayValue() || '');
        m.sh.getRange(row, nc + 1).setValue((old ? old + ' | ' : '') + '⚠ ' + sup + ': ' + String(b.problem).slice(0, 300));
        return { ok: true, problem: true };
      }
      if (reason !== 'Work' && reason !== 'Rework') return { ok: false, error: 'Choose Work or Rework' };
      m.sh.getRange(row, m.C.reason + 1).setValue(reason);
      m.sh.getRange(row, m.C.verifiedBy + 1).setValue(sup);
      m.sh.getRange(row, m.C.verifiedAt + 1).setNumberFormat('@').setValue(llNow_());
      return { ok: true };
    });
  }
  return { ok: false, error: 'unknown' };
}

function llMine_(a, b) {
  var user = String(b.user || '').trim();

  if (a === 'list') return { ok: true, entries: llReadLog_() };

  if (a === 'save') {
    var list = b.entries || (b.entry ? [b.entry] : []);
    for (var i = 0; i < list.length; i++) {
      if (!/^\d{3,}/.test(String((list[i] || {}).offer || '').trim()))
        return { ok: false, error: 'offer', message: 'Offer No is required (must start with a number)' };
    }
    return llLocked_(function () {
      var m = llLogMap_(true);
      list.forEach(function (x) { if (!x.by) x.by = user; llWriteEntry_(m, x, {}); });
      return { ok: true };
    });
  }

  if (a === 'delete') {
    return llLocked_(function () {
      var m = llLogMap_(false), row = llFindRow_(m, b.id);
      if (row) m.sh.deleteRow(row);
      return { ok: true };
    });
  }

  if (a === 'config') {
    var props = PropertiesService.getScriptProperties();
    var cfg = JSON.parse(props.getProperty('LL_CONFIG') || '{}');
    delete cfg.contractors;
    var cs = llReadContr_();
    var old = JSON.parse(props.getProperty('LL_CONFIG') || '{}').contractors;
    if (!cs.length && Array.isArray(old) && old.length) { llWriteContr_(old); cs = llReadContr_(); }
    cfg.contractors = cs;
    return { ok: true, config: cfg, invoices: llReadInv_() };
  }

  if (a === 'setConfig') {
    var src = b.config || {};
    return llLocked_(function () {
      llWriteContr_(Array.isArray(src.contractors) ? src.contractors : []);
      PropertiesService.getScriptProperties().setProperty('LL_CONFIG',
        JSON.stringify({ co: String(src.co || ''), prefix: String(src.prefix || '') }));
      return { ok: true };
    });
  }

  if (a === 'reserveInvoice' || a === 'logInvoice') {
    var inv = b.invoice || {};
    return llLocked_(function () {
      var all = llReadInv_(), no = String(inv.no || '').trim();
      if (a === 'reserveInvoice') {
        // next number for THIS contractor: highest of (its last bill no, its invoices) + 1
        var c = llReadContr_().filter(function (x) { return x.id === String(inv.contractorId); })[0] || {};
        var mine = all.filter(function (r) { return r.contractorId === String(inv.contractorId); }).map(function (r) { return r.no; });
        no = llNextNo_([c.lastBill || ''].concat(mine));
        var taken = {}; all.forEach(function (r) { taken[r.no] = 1; });
        while (taken[no]) no = llNextNo_([no]);
      } else {
        if (!no) return { ok: false, error: 'Invoice no missing' };
        if (all.some(function (r) { return r.no === no; })) return { ok: true, no: no, duplicate: true };
      }
      var sh = llInvSheet_(), row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, LL_INV_HEAD.length).setNumberFormat('@');
      sh.getRange(row, 1, 1, LL_INV_HEAD.length).setValues([[llNow_(), no, inv.date || '', inv.period || '',
        inv.purpose || 'Both', inv.reason || 'Both', inv.filter || '', inv.billedBy || '', inv.billedTo || '',
        String(Number(inv.entries) || 0), String(Number(inv.subtotal) || 0), String(Number(inv.gst) || 0),
        String(Number(inv.amount) || 0), inv.by || user, String(inv.contractorId || '')]]);
      return { ok: true, no: no };
    });
  }

  if (a === 'reqList') {
    var rs = llReadReqs_();
    rs.sort(function (x, y) { return x.raisedAt < y.raisedAt ? 1 : -1; });
    var pend = rs.filter(function (q) { return q.status === 'Pending'; });
    var done = rs.filter(function (q) { return q.status !== 'Pending'; }).slice(0, 60);
    return { ok: true, requests: pend.concat(done) };
  }

  if (a === 'reqApprove') {
    var rate = Number(b.rate) || 0;
    if (rate <= 0) return { ok: false, error: 'Enter the rate' };
    return llLocked_(function () {
      var q = llFindReq_(b.id); if (!q) return { ok: false, error: 'Request not found' };
      if (q.status !== 'Pending') return { ok: false, error: 'Already ' + q.status.toLowerCase() + ' by ' + q.decidedBy };
      var entry = {
        id: 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5),
        date: q.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        offer: q.offer, customer: q.customer, location: q.location,
        labours: Number(b.labours) || Number(q.labours) || 0, rate: rate,
        purpose: q.purpose, reason: q.reason,
        pulleys: b.pulleys != null ? Number(b.pulleys) || 0 : Number(q.pulleys) || 0,
        pulleyDays: b.pulleyDays != null ? Number(b.pulleyDays) || 0 : Number(q.pulleyDays) || 0,
        note: q.note, by: user, requestId: q.id
      };
      llWriteEntry_(llLogMap_(true), entry, {});
      q.status = 'Approved'; q.decidedBy = user; q.decidedAt = llNow_(); q.rate = String(rate); q.entryId = entry.id;
      q.history = (q.history ? q.history + ' | ' : '') + q.decidedAt + ' approved by ' + user + ' @ ₹' + rate;
      llSaveReq_(q);
      return { ok: true, entryId: entry.id };
    });
  }

  if (a === 'reqReject') {
    var why = String(b.reason || '').trim();
    if (!why) return { ok: false, error: 'Give a reason' };
    return llLocked_(function () {
      var q = llFindReq_(b.id); if (!q) return { ok: false, error: 'Request not found' };
      if (q.status !== 'Pending') return { ok: false, error: 'Already ' + q.status.toLowerCase() };
      q.status = 'Rejected'; q.decidedBy = user; q.decidedAt = llNow_(); q.rejectReason = why;
      q.history = (q.history ? q.history + ' | ' : '') + q.decidedAt + ' rejected by ' + user + ': ' + why;
      llSaveReq_(q);
      return { ok: true };
    });
  }
  return { ok: false, error: 'unknown' };
}

/* ============================ LABOUR LOG TAB ============================ */
// [key, heading written if missing, headings it may already have]
var LL_LOG_COLS = [
  ['id', 'ID', ['id']], ['date', 'Date', ['date']], ['period', 'Billing Period', ['billingperiod']],
  ['offer', 'Offer No', ['offerno', 'offer', 'offernumber']], ['customer', 'Customer', ['customer', 'customername']],
  ['location', 'Location', ['location', 'site']], ['labours', 'No of Labours', ['nooflabours', 'labours', 'labour']],
  ['purpose', 'Purpose', ['purpose']], ['reason', 'Reason', ['reason']], ['rate', 'Rate', ['rate', 'rateperlabour']],
  ['pulleyAmt', 'Chain Pulley', ['chainpulley']], ['amount', 'Amount', ['amount']],
  ['note', 'Remarks', ['remarks', 'remark', 'note']], ['savedAt', 'Saved At', ['savedat']],
  ['pulleys', 'Pulleys', ['pulleys']], ['pulleyDays', 'Pulley Days', ['pulleydays']],
  ['by', 'Entered By', ['enteredby']], ['requestId', 'Request ID', ['requestid']],
  ['verifiedBy', 'Verified By', ['verifiedby']], ['verifiedAt', 'Verified At', ['verifiedat']]
];
var LL_TEXT_KEYS = ['id', 'date', 'period', 'offer', 'requestId', 'savedAt', 'verifiedAt'];

function llNorm_(h) { return String(h).toLowerCase().replace(/[^a-z]/g, ''); }

function llLogMap_(ensure) {
  var ss = llSS_(), sh = ss.getSheetByName(LL_LOG_TAB);
  if (!sh) { sh = ss.insertSheet(LL_LOG_TAB); sh.setFrozenRows(1); }
  var width = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, width).getDisplayValues()[0];
  while (head.length && !String(head[head.length - 1]).trim()) head.pop();
  var H = head.map(llNorm_), C = {};
  LL_LOG_COLS.forEach(function (c) {
    for (var i = 0; i < c[2].length; i++) { var k = H.indexOf(c[2][i]); if (k > -1) { C[c[0]] = k; return; } }
    if (ensure) { head.push(c[1]); H.push(llNorm_(c[1])); C[c[0]] = head.length - 1; sh.getRange(1, head.length).setValue(c[1]); }
  });
  return { sh: sh, C: C, width: head.length };
}

function llReadLog_() {
  var m = llLogMap_(false), v = m.sh.getDataRange().getDisplayValues(), out = [];
  var g = function (r, k) { return m.C[k] != null ? String(r[m.C[k]] || '').trim() : ''; };
  for (var r = 1; r < v.length; r++) {
    var id = g(v[r], 'id'); if (!id) continue;
    var by = g(v[r], 'by'), rq = g(v[r], 'requestId'), vb = g(v[r], 'verifiedBy');
    var pd = llNum_(g(v[r], 'pulleyDays')), pn = llNum_(g(v[r], 'pulleys')), pa = llNum_(g(v[r], 'pulleyAmt'));
    if (!pd && pa) { pn = 1; pd = Math.max(1, Math.round(pa / LL_PULLEY_RATE)); }   // older rows: amount only
    var legacy = !by && !rq;   // rows from before verification existed count as checked
    out.push({ id: id, date: g(v[r], 'date'), offer: g(v[r], 'offer'), customer: g(v[r], 'customer'),
      location: g(v[r], 'location'), labours: llNum_(g(v[r], 'labours')), rate: llNum_(g(v[r], 'rate')),
      purpose: g(v[r], 'purpose') || 'Unloading', reason: g(v[r], 'reason') || 'Work', note: g(v[r], 'note'),
      pulleys: pd ? (pn || 1) : 0, pulleyDays: pd, by: by, requestId: rq,
      verified: legacy || !!vb, verifiedBy: vb || (legacy ? 'legacy' : ''), verifiedAt: g(v[r], 'verifiedAt') });
  }
  return out;
}

function llFindRow_(m, id) {
  var n = m.sh.getLastRow(); if (n < 2 || m.C.id == null) return 0;
  var ids = m.sh.getRange(2, m.C.id + 1, n - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === String(id)) return i + 2;
  return 0;
}

function llWriteEntry_(m, x, opts) {
  var row = llFindRow_(m, x.id), isNew = !row;
  if (isNew) row = m.sh.getLastRow() + 1;
  var vals = isNew ? new Array(m.width).fill('') : m.sh.getRange(row, 1, 1, m.width).getValues()[0];
  var cur = function (k) { return m.C[k] != null ? String(vals[m.C[k]] || '') : ''; };
  var set = function (k, val) { if (m.C[k] != null) vals[m.C[k]] = val; };
  var verified = !!cur('verifiedBy');
  var labours = Number(x.labours) || 0, rate = Number(x.rate) || 0;
  var pd = Number(x.pulleyDays) || 0, pn = pd ? Math.max(1, Number(x.pulleys) || 1) : 0;
  var pAmt = pn * pd * LL_PULLEY_RATE, d = String(x.date || '').slice(0, 10);
  set('id', String(x.id)); set('date', d);
  set('period', d ? d.slice(0, 7) + (parseInt(d.slice(8, 10), 10) <= 15 ? '-A' : '-B') : '');
  set('offer', String(x.offer || '').trim()); set('customer', x.customer || ''); set('location', x.location || '');
  set('labours', labours); set('purpose', x.purpose || 'Unloading');
  if (!verified) set('reason', x.reason || 'Work');           // once a supervisor verified it, Reason stays
  set('rate', rate); set('pulleyAmt', pAmt); set('amount', labours * rate + pAmt);
  set('note', x.note || ''); set('savedAt', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'M/d/yyyy HH:mm:ss'));
  set('pulleys', pn); set('pulleyDays', pd);
  if (!cur('by')) set('by', x.by || '');
  if (!cur('requestId') && x.requestId) set('requestId', x.requestId);
  LL_TEXT_KEYS.forEach(function (k) { if (m.C[k] != null) m.sh.getRange(row, m.C[k] + 1).setNumberFormat('@'); });
  m.sh.getRange(row, 1, 1, m.width).setValues([vals]);
}

/* ============================ REQUESTS TAB ============================ */
var LL_REQ_HEAD = ['Request ID', 'Raised At', 'Supervisor', 'Offer No', 'Customer', 'Location', 'Date Needed',
  'Labours', 'Purpose', 'Reason', 'Pulleys', 'Pulley Days', 'Priority', 'Remarks', 'Status', 'Decided By',
  'Decided At', 'Rate', 'Reject Reason', 'Entry ID', 'Resends', 'History', 'Map Link'];
var LL_REQ_KEYS = ['id', 'raisedAt', 'sup', 'offer', 'customer', 'location', 'date', 'labours', 'purpose', 'reason',
  'pulleys', 'pulleyDays', 'priority', 'note', 'status', 'decidedBy', 'decidedAt', 'rate', 'rejectReason',
  'entryId', 'resends', 'history', 'mapLink'];

function llReqSheet_() {
  var ss = llSS_(), sh = ss.getSheetByName(LL_REQ_TAB);
  if (!sh) {
    sh = ss.insertSheet(LL_REQ_TAB); sh.getRange('A:W').setNumberFormat('@');
    sh.getRange(1, 1, 1, LL_REQ_HEAD.length).setValues([LL_REQ_HEAD]); sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 23).getDisplayValue()) !== 'Map Link') {
    sh.getRange(1, 23).setValue('Map Link');
  }
  return sh;
}
function llReadReqs_() {
  var v = llReqSheet_().getDataRange().getDisplayValues(), out = [];
  for (var r = 1; r < v.length; r++) {
    if (!String(v[r][0]).trim()) continue;
    var q = { _row: r + 1 };
    LL_REQ_KEYS.forEach(function (k, i) { q[k] = String(v[r][i] || '').trim(); });
    out.push(q);
  }
  return out;
}
function llFindReq_(id) { return llReadReqs_().filter(function (q) { return q.id === String(id); })[0] || null; }
function llReqRow_(q) { return LL_REQ_KEYS.map(function (k) { return q[k] == null ? '' : String(q[k]); }); }
function llSaveReq_(q) { llReqSheet_().getRange(q._row, 1, 1, LL_REQ_KEYS.length).setValues([llReqRow_(q)]); }
function llCleanReq_(r) {
  var o = {
    offer: String(r.offer || '').trim(), customer: String(r.customer || '').trim(),
    location: String(r.location || '').trim(), date: String(r.date || '').slice(0, 10),
    labours: String(Math.max(0, parseInt(r.labours, 10) || 0)),
    purpose: r.purpose === 'Shifting' ? 'Shifting' : 'Unloading', reason: r.reason === 'Rework' ? 'Rework' : 'Work',
    pulleys: String(Math.max(0, parseInt(r.pulleys, 10) || 0)), pulleyDays: String(Math.max(0, parseInt(r.pulleyDays, 10) || 0)),
    priority: ['Urgent', 'High', 'Normal'].indexOf(r.priority) > -1 ? r.priority : 'Normal',
    note: String(r.note || '').slice(0, 500),
    mapLink: llMapLink_(r.mapLink)
  };
  if (!/^\d{3,}/.test(o.offer)) return { error: 'Offer No is required (must start with a number)' };
  if (Number(o.labours) < 1) return { error: 'Enter no. of labours' };
  return o;
}

// keep only a real web link (first http/https URL found in what was pasted)
function llMapLink_(x) {
  var m = String(x || '').match(/https?:\/\/[^\s<>"']+/i);
  return m ? m[0].slice(0, 500) : '';
}

/* ============================ CONTRACTORS TAB ============================ */
var LL_CONTR_HEAD = ['ID', 'Name', 'Address & Phone', 'GSTIN', 'PAN', 'Aadhaar', 'Account Holder', 'Bank', 'Account No', 'IFSC', 'Last Bill No'];
var LL_CONTR_KEYS = ['id', 'name', 'addr', 'gstin', 'pan', 'aadhaar', 'acName', 'bankName', 'acNo', 'ifsc', 'lastBill'];

function llContrSheet_() {
  var ss = llSS_(), sh = ss.getSheetByName(LL_CONTR_TAB);
  if (!sh) {
    sh = ss.insertSheet(LL_CONTR_TAB); sh.getRange('A:K').setNumberFormat('@');
    sh.getRange(1, 1, 1, LL_CONTR_HEAD.length).setValues([LL_CONTR_HEAD]); sh.setFrozenRows(1);
  }
  return sh;
}
function llReadContr_() {
  var sh = llContrSheet_(), v = sh.getDataRange().getDisplayValues(), out = [];
  for (var r = 1; r < v.length; r++) {
    if (!String(v[r][1]).trim()) continue;
    if (!String(v[r][0]).trim()) { v[r][0] = 'c' + Date.now().toString(36) + r; sh.getRange(r + 1, 1).setValue(v[r][0]); }
    var o = {}; LL_CONTR_KEYS.forEach(function (k, i) { o[k] = String(v[r][i] || '').trim(); });
    out.push(o);
  }
  return out;
}
function llWriteContr_(list) {
  var sh = llContrSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, LL_CONTR_HEAD.length).clearContent();
  var rows = (list || []).filter(function (x) { return x && x.name; }).map(function (x) {
    return LL_CONTR_KEYS.map(function (k) { return String(x[k] || ''); });
  });
  if (rows.length) sh.getRange(2, 1, rows.length, LL_CONTR_HEAD.length).setNumberFormat('@').setValues(rows);
}

/* ============================ INVOICES TAB ============================ */
var LL_INV_HEAD = ['Logged At', 'Invoice No', 'Invoice Date', 'Period', 'Purpose', 'Reason', 'Filter',
  'Billed By', 'Billed To', 'Entries', 'Subtotal', 'GST %', 'Amount', 'Raised By', 'Contractor ID'];

function llInvSheet_() {
  var ss = llSS_(), sh = ss.getSheetByName(LL_INV_TAB);
  if (!sh) {
    sh = ss.insertSheet(LL_INV_TAB); sh.getRange('A:O').setNumberFormat('@');
    sh.getRange(1, 1, 1, LL_INV_HEAD.length).setValues([LL_INV_HEAD]); sh.setFrozenRows(1);
  } else if (String(sh.getRange(1, 15).getDisplayValue()) !== 'Contractor ID') {
    sh.getRange(1, 15).setValue('Contractor ID');
  }
  return sh;
}
function llReadInv_() {
  var v = llInvSheet_().getDataRange().getDisplayValues(); v.shift();
  return v.filter(function (r) { return String(r[1]).trim(); }).map(function (r) {
    return { no: String(r[1]).trim(), date: String(r[2]).slice(0, 10), period: String(r[3]), billedBy: String(r[7] || ''),
      amount: llNum_(r[12]), by: String(r[13] || ''), contractorId: String(r[14] || '').trim() };
  });
}
// "GME/25-26/045" → "GME/25-26/046" · "124" → "125" · picks the highest of the list
function llNextNo_(list) {
  var best = null, bestN = -1;
  (list || []).forEach(function (s) {
    var m = String(s || '').trim().match(/^(.*?)(\d+)(\D*)$/);
    if (m && Number(m[2]) > bestN) { bestN = Number(m[2]); best = m; }
  });
  if (!best) return '1';
  var n = String(bestN + 1); while (n.length < best[2].length) n = '0' + n;
  return best[1] + n + best[3];
}

/* ============================ HELPERS ============================ */
function llSS_() { return LL_SS_ID ? SpreadsheetApp.openById(LL_SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function llNum_(x) { return Number(String(x).replace(/[^0-9.\-]/g, '')) || 0; }
function llNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); }
function llLocked_(fn) { var l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
function llOut_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
