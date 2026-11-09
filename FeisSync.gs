/**
 * Feis Live Results — sync hub for Google Sheets (Peters School of Irish Dance)
 *
 * Paste this into Extensions → Apps Script of a new Google Sheet, then
 * Deploy → New deployment → Web app → Execute as: Me · Who has access: Anyone → Deploy.
 * Copy the Web app URL into the app (Event setup → Sync).
 *
 * The first device that connects sets the event key; every request must carry it.
 * Data rows live on the "sync" tab (don't edit them by hand). The "Results" tab is
 * a readable copy of all results, refreshed as scores come in.
 */
function sh_() { var ss = SpreadsheetApp.getActive(); var s = ss.getSheetByName('sync'); if (!s) { s = ss.insertSheet('sync'); } return s; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function auth_(k) {
  if (!k) return false;
  var p = PropertiesService.getScriptProperties(), cur = p.getProperty('EVENT_KEY');
  if (!cur) { p.setProperty('EVENT_KEY', k); return true; }
  return cur === k;
}
function doGet(e) {
  var q = e.parameter || {};
  if (q.action === 'ping') return out_({ ok: true, now: Date.now() });
  if (!auth_(q.key)) return out_({ error: 'bad key' });
  if (q.action === 'tabs') {   // feis settings tabs, read as shown in the Sheet
    var ss = SpreadsheetApp.getActive(), out = {};
    SETTINGS_TABS.forEach(function (n) { var s = ss.getSheetByName(n); if (s && s.getLastRow()) out[n] = s.getRange(1, 1, s.getLastRow(), Math.max(1, s.getLastColumn())).getDisplayValues(); });
    return out_({ tabs: out });
  }
  if (q.action === 'photo') {   // sheet photo from Drive, as a data URL
    var f = DriveApp.getFileById(q.id), b = f.getBlob();
    return out_({ data: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) });
  }
  var since = Number(q.since || 0), s = sh_(), n = s.getLastRow(), now = Date.now();
  if (n < 1) return out_({ now: now, changes: [] });
  var head = s.getRange(1, 1, n, 3).getValues(), changes = [], w = Math.max(4, s.getLastColumn());
  for (var i = 0; i < head.length; i++) {
    if (Number(head[i][1]) > since) {
      var r = s.getRange(i + 1, 4, 1, w - 3).getValues()[0];
      changes.push({ k: head[i][0], t: Number(head[i][1]), d: head[i][2], v: r.join('') });
    }
  }
  return out_({ now: now, changes: changes });
}
function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  if (!auth_(body.key)) return out_({ error: 'bad key' });
  if (body.action === 'photo') return savePhoto_(body);
  if (body.action === 'writeTabs') return writeTabs_(body.sheets || []);
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var s = sh_(), n = s.getLastRow(), now = Date.now();
    var keys = n ? s.getRange(1, 1, n, 1).getValues().map(function (r) { return r[0]; }) : [];
    (body.changes || []).forEach(function (c) {
      var v = String(c.v), chunks = [];
      for (var i = 0; i < v.length; i += 45000) chunks.push("'" + v.slice(i, i + 45000));   // apostrophe keeps Sheets from reading it as a formula/number
      if (!chunks.length) chunks.push("'");
      var row = [c.k, now, body.device || ''].concat(chunks), idx = keys.indexOf(c.k);
      if (idx < 0) { s.appendRow(row); keys.push(c.k); }
      else { var r = idx + 1; s.getRange(r, 1, 1, s.getMaxColumns()).clearContent(); s.getRange(r, 1, 1, row.length).setValues([row]); }
    });
    if (body.results) writeResults_(body.results);
    return out_({ ok: true, now: now });
  } finally { lock.releaseLock(); }
}
function folder_() {
  var name = 'Feis Live Results — sheet photos', it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function savePhoto_(body) {   // body: { k: 'compId:adjIndex', data: 'data:image/jpeg;base64,...' or '' to remove, label }
  var fo = folder_(), name = String(body.k).replace(/[^\w:-]/g, '_') + '.jpg', old = fo.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  var rec = '';
  if (body.data) {
    var m = String(body.data).match(/^data:([^;]+);base64,(.*)$/);
    var file = fo.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
    if (body.label) file.setDescription(body.label);
    rec = JSON.stringify({ id: file.getId(), at: Date.now() });
  }
  return doPost({ postData: { contents: JSON.stringify({ key: body.key, device: body.device, changes: [{ k: 'ph:' + body.k, v: rec }] }) } });
}
function writeResults_(rows) {
  var ss = SpreadsheetApp.getActive(), r = ss.getSheetByName('Results') || ss.insertSheet('Results');
  r.clearContents();
  var data = [['Competition', 'Place', 'Dancer No', 'Name', 'School', 'Grid total', 'Status']].concat(rows);
  r.getRange(1, 1, data.length, 7).setValues(data);
  r.getRange(1, 1, 1, 7).setFontWeight('bold'); r.setFrozenRows(1);
}

var SETTINGS_TABS = ['Feis details', 'Stages', 'Blocks', 'Dance timing', 'Competitions offered', 'Adjudicators', 'Points'];
function writeTabs_(sheets) {   // feis settings from the app → tabs with headers, widths and drop-down lists
  var ss = SpreadsheetApp.getActive();
  sheets.forEach(function (sh) {
    var s = ss.getSheetByName(sh.name) || ss.insertSheet(sh.name);
    s.clear(); s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).clearDataValidations();
    var w = 1; sh.rows.forEach(function (r) { w = Math.max(w, r.length); });
    var data = sh.rows.map(function (r) { var a = r.slice(); while (a.length < w) a.push(''); return a.map(function (v) { return v === null || v === undefined ? '' : String(v); }); });
    if (data.length) { var rg = s.getRange(1, 1, data.length + 60, w); rg.setNumberFormat('@'); s.getRange(1, 1, data.length, w).setValues(data); }
    if (sh.header && data.length) { s.getRange(1, 1, 1, w).setFontWeight('bold').setFontColor('#ffffff').setBackground('#b5203a'); s.setFrozenRows(1); }
    (sh.widths || []).forEach(function (cw, i) { s.setColumnWidth(i + 1, Math.round(cw * 7.5)); });
    (sh.validations || []).forEach(function (v) { var rule = SpreadsheetApp.newDataValidation().requireValueInList(v.list, true).setAllowInvalid(!v.strict).build(); s.getRange(v.ref).setDataValidation(rule); });
  });
  return out_({ ok: true });
}
