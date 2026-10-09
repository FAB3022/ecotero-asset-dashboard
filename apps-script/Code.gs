/**
 * Asset securement dashboard API (Google Apps Script, bound to the tracker spreadsheet).
 *
 * This file holds no project data. It reads and writes the spreadsheet it is attached to,
 * plus a companion "Project Log" spreadsheet that it creates in the owner's Drive.
 *
 * One-time setup
 *   1. In the tracker: Extensions > Apps Script. Replace the editor contents with this file and save.
 *   2. Pick the function "setup" in the toolbar and press Run. Approve the permission prompt.
 *      The execution log prints the access key and the Project Log link.
 *   3. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone.
 *      Copy the web app URL, then enter that URL and the access key in the dashboard.
 *
 * After editing this file later: Deploy > Manage deployments > Edit > Version: New version > Deploy.
 * To replace the access key: run "rotateKey", then give the team the new key.
 */

var REGISTER_TAB = 'Asset Register';
var ID_HEADER = 'Asset ID';
var ID_PREFIX = 'AST-';
var LOG_TITLE = 'Ecotero Project Log';
var WEEKLY_DAY = 'FRIDAY';
var WEEKLY_HOUR = 9;
var VERSION = 1;

var COL = {
  status: 'Status', brand: 'Brand', owner: 'Action Owner', ownerStatus: 'Owner Status', blocker: 'Blocker / Decision',
  evidence: 'Evidence Link', renewal: 'Renewal / Deadline', notes: 'Historical Notes / Source',
  target: 'Target Name (Unverified)', actual: 'Actual Handle / Account ID', platform: 'Platform / Asset'
};
var DATE_HEADERS = ['Target Date', 'Last Verified', 'Renewal / Deadline'];
var S_VERIFIED = 'Secured - verified', S_NA = 'Not applicable', S_DEFAULT = 'Needs verification';

var LOG = {
  milestones: {tab: 'Milestones', prefix: 'M', pad: 1, fields: [['id', 'ID'], ['title', 'Title'], ['date', 'Date'], ['state', 'State'], ['gate', 'Gate'], ['note', 'Note'], ['updatedAt', 'Updated']]},
  actions: {tab: 'Actions', prefix: 'P', pad: 2, fields: [['id', 'ID'], ['action', 'Action'], ['owner', 'Owner'], ['ownerState', 'Owner Status'], ['status', 'Status'], ['due', 'Due'], ['dependency', 'Depends On'], ['acceptance', 'Acceptance'], ['note', 'Note'], ['updatedAt', 'Updated']]},
  evidence: {tab: 'Evidence', prefix: 'K', pad: 2, fields: [['id', 'ID'], ['claim', 'Claim'], ['type', 'Type'], ['eventDate', 'Event Date'], ['observedDate', 'Observed Date'], ['source', 'Source'], ['state', 'State'], ['supersedes', 'Supersedes'], ['updatedAt', 'Updated']]},
  decisions: {tab: 'Decisions', prefix: 'D', pad: 2, fields: [['id', 'ID'], ['kind', 'Kind'], ['title', 'Title'], ['detail', 'Detail'], ['owner', 'Owner'], ['status', 'Status'], ['decidedOn', 'Decided On'], ['outcome', 'Outcome'], ['updatedAt', 'Updated']]},
  team: {tab: 'Team', prefix: 'T', pad: 2, fields: [['id', 'ID'], ['name', 'Name'], ['role', 'Role'], ['state', 'Role Status'], ['inChannel', 'In Channel'], ['note', 'Note'], ['updatedAt', 'Updated']]},
  updates: {tab: 'Weekly Updates', prefix: 'W', pad: 2, fields: [['id', 'ID'], ['date', 'Date'], ['sent', 'State'], ['summary', 'Summary'], ['changes', 'Changes'], ['blockers', 'Blockers'], ['nextWeek', 'Next Week'], ['forecast', 'Forecast'], ['decisionsNeeded', 'Decisions Needed'], ['updatedAt', 'Updated']]}
};
var AUX = {
  snapshots: {tab: 'Snapshots', headers: ['Date', 'Metrics']},
  changes: {tab: 'Change Log', headers: ['Time', 'Source', 'Asset ID', 'Field', 'Old', 'New']},
  removed: {tab: 'Removed', headers: ['Time', 'List', 'ID', 'Data']}
};

/* ---------- entry points ---------- */

function doGet(){ return json_({ok: true, service: 'asset-dashboard-api', version: VERSION}); }

function doPost(e){
  var out;
  try { out = route_(JSON.parse((e && e.postData && e.postData.contents) || '{}')); }
  catch (err){ out = err_('server_error', String((err && err.message) || err)); }
  return json_(out);
}

function route_(req){
  if (req.action === 'ping') return {ok: true, service: 'asset-dashboard-api', version: VERSION};
  var key = props_().getProperty('ACCESS_KEY');
  if (!key) return err_('not_set_up', 'Run the setup function in the script editor first.');
  if (String(req.key || '') !== key){ Utilities.sleep(700); return err_('unauthorized', 'That access key was not accepted.'); }
  switch (req.action){
    case 'read': return read_();
    case 'saveAsset': return locked_(function(){ return saveAsset_(req); });
    case 'addAsset': return locked_(function(){ return addAsset_(req); });
    case 'deleteAsset': return locked_(function(){ return deleteAsset_(req); });
    case 'saveRecord': return locked_(function(){ return saveRecord_(req); });
    case 'deleteRecord': return locked_(function(){ return deleteRecord_(req); });
    case 'import': return locked_(function(){ return import_(req); });
    case 'draftWeekly': return locked_(function(){ return draftWeekly_(true); });
    default: return err_('unknown_action', 'Unknown action.');
  }
}

/** Run once from the editor. Safe to run again: it keeps the existing key and Project Log. */
function setup(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the tracker: Extensions > Apps Script.');
  if (!ss.getSheetByName(REGISTER_TAB)) throw new Error('This spreadsheet has no "' + REGISTER_TAB + '" tab.');
  var p = props_();
  p.setProperty('TRACKER_ID', ss.getId());
  if (!p.getProperty('ACCESS_KEY')) p.setProperty('ACCESS_KEY', newKey_());
  var log = log_();
  var mine = {weeklyDraft: 1, logSheetEdit: 1};
  ScriptApp.getProjectTriggers().forEach(function(t){ if (mine[t.getHandlerFunction()]) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('weeklyDraft').timeBased().onWeekDay(ScriptApp.WeekDay[WEEKLY_DAY]).atHour(WEEKLY_HOUR).create();
  ScriptApp.newTrigger('logSheetEdit').forSpreadsheet(ss).onEdit().create();
  Logger.log('Setup complete.');
  Logger.log('ACCESS KEY (share privately, never in the sheet): ' + p.getProperty('ACCESS_KEY'));
  Logger.log('Project Log: ' + log.getUrl());
  Logger.log('Next: Deploy > New deployment > Web app > Execute as Me, access Anyone. Copy the web app URL.');
}

/** Replaces the access key. Everyone has to enter the new key in the dashboard. */
function rotateKey(){
  var key = newKey_();
  props_().setProperty('ACCESS_KEY', key);
  Logger.log('NEW ACCESS KEY: ' + key);
}

/** Optional. Adds an "Owner Status" dropdown column (Proposed / Accepted) after "Action Owner". */
function addOwnerStatusColumn(){
  var sh = tracker_().getSheetByName(REGISTER_TAB), headers = headers_(sh);
  if (headers.indexOf(COL.ownerStatus) >= 0){ Logger.log('The column already exists.'); return; }
  var at = headers.indexOf(COL.owner);
  if (at < 0) throw new Error('No "' + COL.owner + '" column to place it after.');
  sh.insertColumnAfter(at + 1);
  sh.getRange(1, at + 2).setValue(COL.ownerStatus);
  var rows = Math.max(sh.getLastRow() - 1, 1);
  sh.getRange(2, at + 2, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Proposed', 'Accepted'], true).setAllowInvalid(false).build());
  Logger.log('Added "' + COL.ownerStatus + '".');
}

/** Weekly trigger. Records a snapshot and writes a draft update if none exists for today. */
function weeklyDraft(){ locked_(function(){ return draftWeekly_(false); }); }

/** Edit trigger. Records changes people make directly in the Asset Register. */
function logSheetEdit(e){
  try {
    var r = e.range, sh = r.getSheet();
    if (sh.getName() !== REGISTER_TAB) return;
    if (r.getNumRows() === 1 && r.getNumColumns() === 1){
      if (r.getRow() === 1) return;
      var headers = headers_(sh), idCol = headers.indexOf(ID_HEADER);
      var id = idCol < 0 ? '' : trim_(sh.getRange(r.getRow(), idCol + 1).getDisplayValue());
      appendChanges_([[stamp_(), 'Sheet', id, headers[r.getColumn() - 1] || '', e.oldValue == null ? '' : String(e.oldValue), e.value == null ? '' : String(e.value)]]);
    } else {
      appendChanges_([[stamp_(), 'Sheet', '', 'Range ' + r.getA1Notation(), '', r.getNumRows() + ' rows by ' + r.getNumColumns() + ' columns edited']]);
    }
  } catch (err){}
}

/* ---------- reading ---------- */

function read_(){
  var reg = readRegister_(), ss = log_(), log = {};
  Object.keys(LOG).forEach(function(c){ log[c] = readTab_(ss, LOG[c]); });
  var handlers = {};
  try { ScriptApp.getProjectTriggers().forEach(function(t){ handlers[t.getHandlerFunction()] = 1; }); } catch (err){}
  var body = {
    register: {headers: reg.headers, rows: reg.rows, options: reg.options, dateCols: reg.dateCols, url: reg.url, name: reg.name},
    log: log,
    snapshots: readSnapshots_(ss).slice(-26),
    changes: readChanges_(ss, 200),
    logUrl: ss.getUrl(),
    triggers: {weekly: !!handlers.weeklyDraft, edits: !!handlers.logSheetEdit}
  };
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(body), Utilities.Charset.UTF_8);
  body.rev = Utilities.base64Encode(digest);
  body.ok = true;
  body.version = VERSION;
  body.readAt = stamp_();
  return body;
}

function readRegister_(){
  var ss = tracker_(), sh = ss.getSheetByName(REGISTER_TAB);
  if (!sh) throw new Error('The "' + REGISTER_TAB + '" tab was not found.');
  var range = sh.getDataRange(), vals = range.getValues(), disp = range.getDisplayValues(), tz = ss.getSpreadsheetTimeZone();
  var headers = disp[0].map(trim_), n = headers.length;
  while (n > 0 && !headers[n - 1]) n--;
  headers = headers.slice(0, n);
  var rows = [];
  for (var r = 1; r < vals.length; r++){
    var row = [], any = false;
    for (var c = 0; c < n; c++){
      var v = vals[r][c], s = (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : trim_(disp[r][c]);
      if (s) any = true;
      row.push(s);
    }
    if (any) rows.push(row);
  }
  var options = {}, dateCols = [];
  if (sh.getLastRow() >= 2 && n){
    var rules = sh.getRange(2, 1, 1, n).getDataValidations()[0];
    for (var i = 0; i < n; i++){
      if (!rules[i]) continue;
      var type = String(rules[i].getCriteriaType()), cv = rules[i].getCriteriaValues();
      if (type === 'VALUE_IN_LIST' && cv && cv[0] && cv[0].length) options[headers[i]] = cv[0].map(String);
      else if (type.indexOf('DATE') === 0) dateCols.push(headers[i]);
    }
  }
  DATE_HEADERS.forEach(function(h){ if (headers.indexOf(h) >= 0 && dateCols.indexOf(h) < 0) dateCols.push(h); });
  return {headers: headers, rows: rows, options: options, dateCols: dateCols, tz: tz, url: ss.getUrl(), name: ss.getName()};
}

function readTab_(ss, def){
  var sh = ss.getSheetByName(def.tab);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getDataRange().getDisplayValues(), hdr = vals[0].map(trim_), out = [];
  for (var r = 1; r < vals.length; r++){
    var o = {};
    def.fields.forEach(function(f){ var i = hdr.indexOf(f[1]); o[f[0]] = i < 0 ? '' : String(vals[r][i]); });
    o.id = trim_(o.id);
    if (o.id) out.push(o);
  }
  return out;
}

function readSnapshots_(ss){
  var sh = ss.getSheetByName(AUX.snapshots.tab);
  if (!sh || sh.getLastRow() < 2) return [];
  var out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function(r){
    try { if (r[0]) out.push({date: trim_(r[0]), metrics: JSON.parse(r[1])}); } catch (err){}
  });
  out.sort(function(a, b){ return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return out;
}

function readChanges_(ss, limit){
  var sh = ss.getSheetByName(AUX.changes.tab);
  if (!sh || sh.getLastRow() < 2) return [];
  var last = sh.getLastRow(), n = Math.min(limit, last - 1);
  return sh.getRange(last - n + 1, 1, n, 6).getDisplayValues().reverse().map(function(r){
    return {time: r[0], source: r[1], id: r[2], field: r[3], old: r[4], 'new': r[5]};
  });
}

/* ---------- asset register writes ---------- */

function saveAsset_(req){
  var sh = tracker_().getSheetByName(REGISTER_TAB), headers = headers_(sh), idCol = headers.indexOf(ID_HEADER), id = trim_(req.id);
  if (idCol < 0) return err_('no_id_column', 'The sheet has no "' + ID_HEADER + '" column.');
  var row = findRow_(sh, idCol, id);
  if (row < 0) return err_('not_found', id + ' is no longer in the sheet.');
  var changes = req.changes || {}, tz = tracker_().getSpreadsheetTimeZone(), cells = [], conflicts = [], bad = [];
  Object.keys(changes).forEach(function(h){
    var c = headers.indexOf(h);
    if (c < 0 || h === ID_HEADER) return;
    var cell = sh.getRange(row, c + 1), cur = cellText_(cell, tz), next = trim_(changes[h]['new']), old = changes[h].old;
    if (next.charAt(0) === '=') bad.push(h);
    if (old != null && cur !== trim_(old) && cur !== next) conflicts.push(h);
    cells.push({h: h, cell: cell, cur: cur, next: next});
  });
  if (bad.length) return err_('invalid_value', 'A value cannot start with "=".', {fields: bad});
  if (conflicts.length) return err_('conflict', 'Changed in the sheet since you opened it.', {fields: conflicts});
  var log = [], failed = [];
  cells.forEach(function(x){
    if (x.cur === x.next) return;
    try { x.cell.setValue(x.next); log.push([stamp_(), 'Dashboard', id, x.h, x.cur, x.next]); }
    catch (err){ failed.push(x.h); }
  });
  SpreadsheetApp.flush();
  appendChanges_(log);
  if (failed.length) return err_('rejected', 'The sheet did not accept a value.', {fields: failed});
  return {ok: true, id: id, changed: log.length};
}

function addAsset_(req){
  var sh = tracker_().getSheetByName(REGISTER_TAB), headers = headers_(sh), idCol = headers.indexOf(ID_HEADER);
  if (idCol < 0) return err_('no_id_column', 'The sheet has no "' + ID_HEADER + '" column.');
  var last = sh.getLastRow(), n = headers.length;
  if (last < 2) return err_('empty_register', 'Add the first row in the sheet, then use the dashboard.');
  var ids = sh.getRange(2, idCol + 1, last - 1, 1).getDisplayValues().map(function(r){ return trim_(r[0]); });
  var vals = req.values || {}, id = trim_(vals[ID_HEADER]);
  if (!id){
    var max = 0;
    ids.forEach(function(s){ if (s.indexOf(ID_PREFIX) === 0){ var k = parseInt(s.slice(ID_PREFIX.length), 10); if (k > max) max = k; } });
    id = ID_PREFIX + ('0000' + (max + 1)).slice(-4);
  }
  if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,39}$/.test(id)) return err_('invalid_id', 'Use letters, numbers, hyphens or underscores for the ID.');
  if (ids.indexOf(id) >= 0) return err_('duplicate_id', id + ' is already in the sheet.');
  var out = headers.map(function(h){ return h === ID_HEADER ? id : trim_(vals[h]); });
  for (var i = 0; i < out.length; i++) if (out[i].charAt(0) === '=') return err_('invalid_value', 'A value cannot start with "=".', {fields: [headers[i]]});
  var lastData = 1;
  for (var j = ids.length - 1; j >= 0; j--) if (ids[j]){ lastData = j + 2; break; }
  // Insert inside the table so the table range grows with it. The new row lands just above the last row.
  sh.insertRowBefore(lastData);
  try { sh.getRange(lastData, 1, 1, n).setValues([out]); }
  catch (err){
    sh.deleteRow(lastData);
    return err_('rejected', 'The sheet did not accept one of the values: ' + String((err && err.message) || err));
  }
  SpreadsheetApp.flush();
  appendChanges_([[stamp_(), 'Dashboard', id, 'Row added', '', '']]);
  return {ok: true, id: id};
}

function deleteAsset_(req){
  var sh = tracker_().getSheetByName(REGISTER_TAB), headers = headers_(sh), idCol = headers.indexOf(ID_HEADER), id = trim_(req.id);
  var row = idCol < 0 ? -1 : findRow_(sh, idCol, id);
  if (row < 0) return err_('not_found', id + ' is no longer in the sheet.');
  var vals = sh.getRange(row, 1, 1, headers.length).getDisplayValues()[0], data = {};
  headers.forEach(function(h, i){ if (h) data[h] = vals[i]; });
  appendAux_('removed', [[stamp_(), 'assets', id, JSON.stringify(data)]]);
  sh.deleteRow(row);
  appendChanges_([[stamp_(), 'Dashboard', id, 'Row removed', '', '']]);
  return {ok: true, id: id};
}

/* ---------- project log writes ---------- */

function saveRecord_(req){
  var def = LOG[req.collection];
  if (!def) return err_('unknown_list', 'Unknown list.');
  var sh = ensureTab_(log_(), def.tab, def.fields.map(function(f){ return f[1]; }));
  var vals = sh.getDataRange().getDisplayValues(), hdr = vals[0].map(trim_), idIdx = hdr.indexOf('ID');
  var v = req.values || {}, id = trim_(req.id || v.id), row = -1;
  for (var r = 1; r < vals.length; r++) if (id && trim_(vals[r][idIdx]) === id){ row = r + 1; break; }
  if (req.create){
    if (row > 0) return err_('duplicate_id', id + ' is already in use.');
    if (!id) id = nextId_(vals, idIdx, trim_(req.prefix) || def.prefix, def.pad);
    if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,39}$/.test(id)) return err_('invalid_id', 'Use letters, numbers, hyphens or underscores for the ID.');
    row = Math.max(sh.getLastRow(), 1) + 1;
  } else if (row < 0) return err_('not_found', 'That record is no longer in the project log.');
  var cur = row <= vals.length ? vals[row - 1] : [], today = today_();
  var out = hdr.map(function(h, i){
    var f = null;
    def.fields.forEach(function(x){ if (x[1] === h) f = x; });
    var keep = cur[i] == null ? '' : cur[i];
    if (!f) return keep;
    if (f[0] === 'id') return id;
    if (f[0] === 'updatedAt') return today;
    return v[f[0]] == null ? keep : String(v[f[0]]);
  });
  sh.getRange(row, 1, 1, hdr.length).setNumberFormat('@').setValues([out]);
  return {ok: true, id: id};
}

function deleteRecord_(req){
  var def = LOG[req.collection];
  if (!def) return err_('unknown_list', 'Unknown list.');
  var sh = log_().getSheetByName(def.tab), id = trim_(req.id);
  if (!sh) return err_('not_found', 'That record is no longer in the project log.');
  var vals = sh.getDataRange().getDisplayValues(), hdr = vals[0].map(trim_), idIdx = hdr.indexOf('ID');
  for (var r = 1; r < vals.length; r++){
    if (trim_(vals[r][idIdx]) !== id) continue;
    var data = {};
    hdr.forEach(function(h, i){ if (h) data[h] = vals[r][i]; });
    appendAux_('removed', [[stamp_(), req.collection, id, JSON.stringify(data)]]);
    sh.deleteRow(r + 1);
    return {ok: true, id: id};
  }
  return err_('not_found', 'That record is no longer in the project log.');
}

/** Adds records whose IDs are not in the log yet. Existing records are left alone. */
function import_(req){
  var data = req.data || {}, added = {}, skipped = {};
  Object.keys(LOG).forEach(function(c){
    var list = data[c];
    if (!list || !list.length) return;
    var have = {};
    readTab_(log_(), LOG[c]).forEach(function(o){ have[o.id] = 1; });
    added[c] = 0; skipped[c] = 0;
    list.forEach(function(rec){
      var id = trim_(rec && rec.id);
      if (!id || have[id]){ skipped[c]++; return; }
      var res = saveRecord_({collection: c, id: id, create: true, values: rec});
      if (res.ok){ added[c]++; have[id] = 1; } else skipped[c]++;
    });
  });
  return {ok: true, added: added, skipped: skipped};
}

/* ---------- weekly summary ---------- */

function metrics_(reg){
  var at = {};
  reg.headers.forEach(function(h, i){ at[h] = i; });
  var get = function(row, h){ return at[h] == null ? '' : row[at[h]]; };
  var by = {}, total = 0, na = 0, named = 0, accepted = 0, evidence = 0, blocked = 0, brands = {};
  reg.rows.forEach(function(row){
    if (!get(row, ID_HEADER)) return;
    total++;
    var s = get(row, COL.status) || S_DEFAULT;
    by[s] = (by[s] || 0) + 1;
    if (s === S_NA){ na++; return; }
    brands[get(row, COL.brand) || '(none)'] = 1;
    if (get(row, COL.owner)){ named++; if (get(row, COL.ownerStatus) === 'Accepted') accepted++; }
    if (get(row, COL.evidence)) evidence++;
    if (get(row, COL.blocker)) blocked++;
  });
  return {total: total, applicable: total - na, na: na, by: by, verified: by[S_VERIFIED] || 0, named: named,
    accepted: at[COL.ownerStatus] == null ? null : accepted, evidence: evidence, blocked: blocked, brands: Object.keys(brands).length};
}

function draftWeekly_(manual){
  var reg = readRegister_(), m = metrics_(reg), ss = log_(), today = today_();
  var snaps = readSnapshots_(ss), prev = null;
  for (var i = snaps.length - 1; i >= 0; i--) if (snaps[i].date < today){ prev = snaps[i]; break; }
  upsertSnapshot_(ss, today, m);
  var mine = null;
  readTab_(ss, LOG.updates).forEach(function(u){ if (u.date === today) mine = u; });
  if (mine && (mine.sent === 'Sent' || !manual)) return {ok: true, id: mine.id, created: false, kept: true};
  var t = composeWeekly_(reg, m, prev, ss, today);
  var res = saveRecord_({collection: 'updates', id: mine ? mine.id : '', create: !mine,
    values: {date: today, sent: 'Draft', summary: t.summary, changes: t.changes, blockers: t.blockers, nextWeek: t.nextWeek, forecast: t.forecast, decisionsNeeded: t.decisions}});
  if (!res.ok) return res;
  return {ok: true, id: res.id, created: !mine};
}

function composeWeekly_(reg, m, prev, ss, today){
  var p = prev ? prev.metrics : null, since = prev ? ' since ' + nice_(prev.date) : '';
  var basis = prev ? 'Changes in brackets are measured against the snapshot of ' + nice_(prev.date) + '.' : 'This is the first snapshot, so there is nothing to compare against yet.';
  var delta = function(cur, old){
    if (!p || old == null) return '';
    var x = cur - old;
    return ' (' + (x > 0 ? '+' + x : x < 0 ? String(x) : 'no change') + ')';
  };
  var st = function(name){ return m.by[name] || 0; };
  var was = function(name){ return p && p.by ? (p.by[name] || 0) : null; };
  var summary = [
    'Applicable scope: ' + m.applicable + ' assets across ' + m.brands + ' brands' + (m.na ? ' (' + m.na + ' not applicable, excluded)' : '') + '.',
    'Secured and verified: ' + m.verified + ' of ' + m.applicable + delta(m.verified, p && p.verified) + '.',
    'Reported, audit pending: ' + st('Reported secured - audit pending') + delta(st('Reported secured - audit pending'), was('Reported secured - audit pending')) + '. In progress: ' + st('In progress') + delta(st('In progress'), was('In progress')) + '.',
    'Awaiting approval: ' + st('Available - approval needed') + '. Blocked or third-party held: ' + (st('Blocked') + st('Third-party held - review')) + '. Needs verification: ' + st(S_DEFAULT) + delta(st(S_DEFAULT), was(S_DEFAULT)) + '.',
    m.accepted == null ? 'Owners named: ' + m.named + ' of ' + m.applicable + delta(m.named, p && p.named) + '.' : 'Owners: ' + m.accepted + ' accepted, ' + (m.named - m.accepted) + ' proposed, ' + (m.applicable - m.named) + ' unassigned.',
    'Evidence links recorded: ' + m.evidence + ' of ' + m.applicable + delta(m.evidence, p && p.evidence) + '.',
    basis
  ].join('\n');

  var from = prev ? prev.date : '', status = [], other = 0, touched = {};
  readChanges_(ss, 1000).forEach(function(c){
    if (from && String(c.time).slice(0, 10) <= from) return;
    if (c.field === COL.status && c.id) status.push(c.id + ': ' + (c.old || 'blank') + ' to ' + (c['new'] || 'blank'));
    else { other++; if (c.id) touched[c.id] = 1; }
  });
  var changes = status.length ? status.slice(0, 15).reverse().join('\n') + (status.length > 15 ? '\nand ' + (status.length - 15) + ' more status changes.' : '') : 'No status changes recorded' + (prev ? since : '') + '.';
  if (other) changes += '\n' + other + ' other edits across ' + Object.keys(touched).length + ' assets.';

  var at = {};
  reg.headers.forEach(function(h, i){ at[h] = i; });
  var get = function(row, h){ return at[h] == null ? '' : row[at[h]]; }, soon = [], limit = addDays_(today, 60);
  reg.rows.forEach(function(row){
    var d = get(row, COL.renewal), confirmed = true, hit;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)){ hit = /renewal date (\d{4}-\d{2}-\d{2})/i.exec(get(row, COL.notes)); d = hit ? hit[1] : ''; confirmed = false; }
    if (d && d <= limit) soon.push({d: d, text: nice_(d) + ' ' + (get(row, COL.actual) || get(row, COL.target) || get(row, COL.platform)) + ' (' + (confirmed ? 'confirmed' : 'historical date, unverified') + ')'});
  });
  soon.sort(function(a, b){ return a.d < b.d ? -1 : 1; });
  var blockers = (m.blocked ? m.blocked + ' register rows carry a blocker or pending decision.' : 'No blockers recorded in the register.') +
    (soon.length ? '\nRenewals within 60 days: ' + soon.slice(0, 6).map(function(x){ return x.text; }).join('; ') + '.' : '');

  var next = readTab_(ss, LOG.actions).filter(function(a){ return a.status !== 'Done'; }).slice(0, 5).map(function(a){
    return '- ' + a.id + ' ' + a.action + (a.owner ? ' (' + a.owner + ', ' + (a.ownerState || 'Proposed').toLowerCase() + ')' : '');
  }).join('\n');
  var ms = readTab_(ss, LOG.milestones).filter(function(x){ return x.state !== 'Done'; }).sort(function(a, b){ return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; })[0];
  var decisions = readTab_(ss, LOG.decisions).filter(function(d){ return d.kind !== 'Risk' && d.status === 'Open'; }).map(function(d){ return '- ' + d.title; }).join('\n');
  return {summary: summary, changes: changes, blockers: blockers, nextWeek: next,
    forecast: ms ? ms.id + ' ' + ms.title + ': ' + nice_(ms.date) + ' (' + (ms.state || 'Proposed').toLowerCase() + ')' : '', decisions: decisions};
}

function upsertSnapshot_(ss, date, metrics){
  var sh = ensureTab_(ss, AUX.snapshots.tab, AUX.snapshots.headers), last = sh.getLastRow(), row = last + 1;
  if (last >= 2){
    var dates = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < dates.length; i++) if (trim_(dates[i][0]) === date){ row = i + 2; break; }
  }
  sh.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([[date, JSON.stringify(metrics)]]);
}

/* ---------- plumbing ---------- */

function props_(){ return PropertiesService.getScriptProperties(); }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function err_(code, message, extra){
  var o = {ok: false, error: code, message: message || ''};
  if (extra) Object.keys(extra).forEach(function(k){ o[k] = extra[k]; });
  return o;
}
function trim_(v){ return v == null ? '' : String(v).trim(); }
function locked_(fn){
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function newKey_(){ return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 32); }
function tracker_(){
  var id = props_().getProperty('TRACKER_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function log_(){
  var p = props_(), id = p.getProperty('LOG_ID'), ss = null;
  if (id){ try { ss = SpreadsheetApp.openById(id); } catch (err){ ss = null; } }
  if (!ss){
    ss = SpreadsheetApp.create(LOG_TITLE);
    p.setProperty('LOG_ID', ss.getId());
  }
  Object.keys(LOG).forEach(function(c){ ensureTab_(ss, LOG[c].tab, LOG[c].fields.map(function(f){ return f[1]; })); });
  Object.keys(AUX).forEach(function(k){ ensureTab_(ss, AUX[k].tab, AUX[k].headers); });
  var blank = ss.getSheetByName('Sheet1');
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
  return ss;
}
function ensureTab_(ss, name, headers){
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function headers_(sh){
  var n = sh.getLastColumn();
  return n ? sh.getRange(1, 1, 1, n).getDisplayValues()[0].map(trim_) : [];
}
function findRow_(sh, idCol, id){
  var last = sh.getLastRow();
  if (!id || last < 2) return -1;
  var ids = sh.getRange(2, idCol + 1, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) if (trim_(ids[i][0]) === id) return i + 2;
  return -1;
}
function cellText_(cell, tz){
  var v = cell.getValue();
  return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : trim_(cell.getDisplayValue());
}
function nextId_(vals, idIdx, prefix, pad){
  var max = 0;
  for (var r = 1; r < vals.length; r++){
    var s = trim_(vals[r][idIdx]);
    if (s.indexOf(prefix) === 0){ var k = parseInt(s.slice(prefix.length), 10); if (k > max) max = k; }
  }
  var num = String(max + 1);
  while (num.length < pad) num = '0' + num;
  return prefix + num;
}
function appendAux_(kind, rows){
  if (!rows.length) return;
  var sh = ensureTab_(log_(), AUX[kind].tab, AUX[kind].headers);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows);
  return sh;
}
function appendChanges_(rows){
  if (!rows.length) return;
  var sh = appendAux_('changes', rows), last = sh.getLastRow();
  if (last > 3000) sh.deleteRows(2, last - 2500);
}
function tz_(){ return tracker_().getSpreadsheetTimeZone(); }
function today_(){ return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function stamp_(){ return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
function addDays_(iso, n){
  var d = new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) + n * 86400000);
  return d.toISOString().slice(0, 10);
}
function nice_(iso){
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  if (!m) return String(iso || '');
  return (+m[3]) + ' ' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][+m[2] - 1] + ' ' + m[1];
}
