/**
 * Paisa — Google Sheets backend (Google Apps Script)
 *
 * Paste this whole file into Extensions → Apps Script of your Google Sheet, then follow SETUP.md.
 * - Data lives in readable tabs: Transactions, Accounts, Snapshots, EMIs, Recurring, Categories, Settings.
 * - The password is never stored in plain text: only a salted SHA-256 hash in Script Properties.
 * - Every request (load/save) must carry the correct password; repeated failures are throttled.
 */

const VERSION = 5; // v2: credit-card tracking (cardId / transfers / account card fields / snapshot update times)

// Column schema per tab. Types: s=text, d=date(yyyy-MM-dd), m=month(yyyy-MM), n=amount, i=integer, a=list, b=boolean
const SCHEMA = {
  Transactions: [['id', 's'], ['date', 'd'], ['type', 's'], ['category', 's'], ['amount', 'n'], ['method', 's'], ['note', 's'], ['tags', 'a'], ['emiId', 's'], ['recurId', 's'], ['cardId', 's'], ['fromId', 's'], ['toId', 's'], ['ts', 'i']],
  Accounts: [['id', 's'], ['name', 's'], ['kind', 's'], ['group', 's'], ['limit', 'n'], ['dueDay', 'i'], ['opening', 'n'], ['openingDate', 'd'], ['openingTs', 'i']],
  Snapshots: [['month', 'm'], ['accountId', 's'], ['account', 's'], ['value', 'n'], ['updated', 'i']],
  EMIs: [['id', 's'], ['name', 's'], ['icon', 's'], ['amount', 'n'], ['tenure', 'i'], ['start', 'm'], ['dueDay', 'i'], ['method', 's'], ['lender', 's'], ['prepaid', 'i']],
  Recurring: [['id', 's'], ['type', 's'], ['name', 's'], ['icon', 's'], ['amount', 'n'], ['category', 's'], ['dueDay', 'i'], ['method', 's'], ['tags', 'a'], ['active', 'b'], ['since', 'm']],
  Categories: [['type', 's'], ['name', 's'], ['icon', 's'], ['color', 's']],
  Settings: [['key', 's'], ['value', 's']],
  Statements: [['id', 's'], ['date', 'd'], ['bank', 's'], ['card', 's'], ['amount', 'n'], ['minDue', 'n'], ['dueDate', 'd'], ['name', 's'], ['fileId', 's'], ['url', 's'], ['ts', 'i'], ['pwHint', 's'], ['product', 's'], ['hash', 's'], ['hidden', 's']],
  Inbox: [['id', 's'], ['date', 'd'], ['type', 's'], ['amount', 'n'], ['merchant', 's'], ['bank', 's'], ['acct', 's'], ['method', 's'], ['snippet', 's'], ['status', 's'], ['txId', 's'], ['ts', 'i']],
};
const MAX_FAILS = 8;          // wrong passwords allowed…
const LOCKOUT_SEC = 15 * 60;  // …before a 15-minute lockout

/* ======================= Web app endpoints ======================= */
function doGet() {
  return json({ ok: true, app: 'Paisa', ver: VERSION, message: 'Backend is running. The app talks to it with POST requests.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, error: 'Bad request' }); }
  try {
    checkPassword(req.password);
    switch (req.action) {
      case 'ping': return json({ ok: true });
      case 'load': return loadLocked();
      case 'statements': return json({ ok: true, statements: readStatements() });
      case 'stmtSave': return json(stmtSave(req.s));
      case 'inbox': return json({ ok: true, inbox: readInbox() });
      case 'inboxMark': return json(inboxMark(req.marks));
      case 'save': return json(saveAll(req.data, req.baseRev, req.deviceId, req.deviceName));
      case 'registerDevice': return json(registerDevice(req));
      case 'unregisterDevice': return json(unregisterDevice(req.deviceId));
      default: return json({ ok: false, error: 'Unknown action' });
    }
  } catch (x) {
    return json({ ok: false, error: x.message, code: x.code || '' });
  }
}

/* ======================= Security ======================= */
function checkPassword(pw) {
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('PW_HASH'), salt = props.getProperty('PW_SALT');
  if (!hash || !salt) throw fail('Password not set yet. In the Google Sheet use menu Paisa → Set / change password.', 'NO_PASSWORD');
  const cache = CacheService.getScriptCache(), fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) throw fail('Too many wrong attempts. Try again in 15 minutes.', 'LOCKED');
  if (sha256(String(pw || '') + salt) !== hash) {
    cache.put('fails', String(fails + 1), LOCKOUT_SEC);
    Utilities.sleep(700);
    throw fail('Incorrect password. Please try again.', 'AUTH');
  }
  if (fails) cache.remove('fails');
}

function sha256(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/* ======================= Sheet menu (owner only) ======================= */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Paisa')
    .addItem('⚡ Set up everything', 'setupEverything')
    .addItem('Set / change password', 'setPasswordPrompt')
    .addSeparator()
    .addItem('Enable bank email import', 'installBankImport')
    .addItem('Import bank emails now', 'importBankEmailsNow')
    .addItem('Import card statements now', 'importStatementsNow')
    .addItem('Disable bank email import', 'removeBankImport')
    .addSeparator()
    .addItem('Unlock after too many attempts', 'clearLockout')
    .addSubMenu(ui.createMenu('Advanced')
      .addItem('Create / repair tabs', 'setupSheets')
      .addItem('Set bank alert sender emails', 'setBankSendersPrompt')
      .addItem('Change bank email search (advanced)', 'setBankQueryPrompt')
      .addSeparator()
      .addItem('Set up push notifications (Firebase key)', 'setFcmKeyPrompt')
      .addItem('Send test notification', 'sendTestPush')
      .addItem('Enable due & daily reminders', 'installReminders')
      .addItem('Disable reminders', 'removeReminders'))
    .addToUi();
  if (!PropertiesService.getScriptProperties().getProperty('PW_HASH')) {
    SpreadsheetApp.getActive().toast('Open the Paisa menu → ⚡ Set up everything to finish setup.', 'Paisa', 10);
  }
}

// One-click first-time setup: password, tabs and (optionally) Gmail bank-alert import.
// Push notifications and reminders are optional extras under Paisa → Advanced.
function setupEverything() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  const had = !!props.getProperty('PW_HASH');
  if (!had || ui.alert('Password already set', 'Keep your current password?\n\nYes = keep it, No = choose a new one.', ui.ButtonSet.YES_NO) === ui.Button.NO) {
    const r = ui.prompt('Step 1 of 3: Choose your Paisa password', 'You will type this in the app to unlock it (min 6 characters). Write it down: it cannot be recovered, only reset.', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    const pw = r.getResponseText();
    if (pw.length < 6) { ui.alert('Password must be at least 6 characters. Run ⚡ Set up everything again.'); return; }
    const salt = Utilities.getUuid();
    props.setProperties({ PW_SALT: salt, PW_HASH: sha256(pw + salt) });
    clearLockout();
  }
  setupSheets();
  let bank = false;
  const ask = ui.alert('Step 3 of 3: Bank email import (optional)', 'Automatically find transactions from your bank / card alert emails in Gmail and list them in the app for you to review?\n\nGoogle will ask you to allow Gmail and Drive access. You can turn this off any time from the Paisa menu.', ui.ButtonSet.YES_NO);
  if (ask === ui.Button.YES) { installBankImport(); bank = true; }
  ui.alert('Setup complete ✓', 'Password saved, tabs created' + (bank ? ', bank email import on' : '') + '.\n\nLast step: in Apps Script click Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone), copy the Web app URL and paste it into the Paisa app.', ui.ButtonSet.OK);
}

function setPasswordPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Set Paisa password', 'Choose a password (min 6 characters). You will type this in the app to unlock it.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText();
  if (pw.length < 6) { ui.alert('Password must be at least 6 characters.'); return; }
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ PW_SALT: salt, PW_HASH: sha256(pw + salt) });
  clearLockout();
  ui.alert('Password saved ✓\n\nOnly a salted hash is stored — the password itself is not saved anywhere.');
}

function clearLockout() { CacheService.getScriptCache().remove('fails'); }

function setupSheets() {
  Object.keys(SCHEMA).forEach(name => {
    const sh = tab(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name].map(c => c[0])]);
    styleHeader(sh, SCHEMA[name].length);
  });
  const ss = SpreadsheetApp.getActive(), def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

// Any manual edit in the sheet bumps the revision so open apps reload instead of overwriting it.
function onEdit() { bumpRev(); }

/* ======================= Read ======================= */
function readAll() {
  const tz = Session.getScriptTimeZone();
  const txns = readTab('Transactions', tz).filter(t => t.date && t.amount).map(t => {
    t.id = t.id || Utilities.getUuid();
    const ty = String(t.type || 'expense').toLowerCase();
    t.type = ty === 'income' || ty === 'transfer' ? ty : 'expense';
    ['emiId', 'recurId', 'cardId', 'fromId', 'toId'].forEach(k => { if (!t[k]) delete t[k]; });
    return t;
  });
  const snaps = {};
  readTab('Snapshots', tz).forEach(r => {
    if (!r.month || !r.accountId) return;
    const s = snaps[r.month] = snaps[r.month] || { values: {}, updated: {} };
    s.values[r.accountId] = r.value;
    if (r.updated) s.updated[r.accountId] = r.updated;
  });
  const catRows = readTab('Categories', tz).filter(c => c.name);
  const settings = {};
  readTab('Settings', tz).forEach(r => { if (r.key) settings[r.key] = parseVal(r.value); });

  const data = {
    txns: txns,
    accounts: readTab('Accounts', tz).filter(a => a.id && a.name).map(a => {
      ['limit', 'dueDay', 'opening', 'openingTs'].forEach(k => { if (!a[k]) delete a[k]; });
      if (!a.openingDate) { delete a.openingDate; delete a.opening; delete a.openingTs; }
      return a;
    }),
    snapshots: Object.keys(snaps).sort().map(m => ({ month: m, values: snaps[m].values, updated: snaps[m].updated })),
    emis: readTab('EMIs', tz).filter(e => e.id && e.name),
    recurring: readTab('Recurring', tz).filter(r => r.id && r.name),
    settings: settings,
  };
  if (catRows.length) data.categories = {
    expense: catRows.filter(c => c.type !== 'income').map(c => ({ name: c.name, icon: c.icon, color: c.color || '#94a3b8' })),
    income: catRows.filter(c => c.type === 'income').map(c => ({ name: c.name, icon: c.icon, color: c.color || '#94a3b8' })),
  };
  return data;
}

function readTab(name, tz) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const head = vals[0].map(h => String(h).trim());
  return vals.slice(1).filter(r => r.some(c => c !== '' && c !== null)).map(r => {
    const o = {};
    SCHEMA[name].forEach(([k, t]) => { const i = head.indexOf(k); if (i >= 0) o[k] = fromCell(r[i], t, tz); });
    return o;
  });
}

function fromCell(v, t, tz) {
  switch (t) {
    case 'n': case 'i': return v === '' ? 0 : Number(String(v).replace(/[,₹\s]/g, '')) || 0;
    case 'd': return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim();
    case 'm': return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM') : String(v).trim().slice(0, 7);
    case 'a': return String(v).split(',').map(s => s.trim()).filter(Boolean);
    case 'b': return !(v === false || String(v).toUpperCase() === 'FALSE');
    default: return String(v);
  }
}

function parseVal(v) { try { return JSON.parse(v); } catch (x) { return v; } }

/* ======================= Write ======================= */
// Reads wait for any in-progress save, so a half-written (empty) sheet is never returned.
function loadLocked() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return json({ ok: true, ver: VERSION, data: readAll(), rev: getRev(), sheetUrl: SpreadsheetApp.getActive().getUrl(), inbox: readInbox(), statements: readStatements() }); } finally { lock.releaseLock(); }
}

function backupSheet() { // at most one copy per day, newest 7 kept, in Drive folder "Paisa Backups"
  try {
    const props = PropertiesService.getScriptProperties(), day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (props.getProperty('LAST_BACKUP') === day) return;
    const it = DriveApp.getFoldersByName('Paisa Backups'), folder = it.hasNext() ? it.next() : DriveApp.createFolder('Paisa Backups');
    const ss = SpreadsheetApp.getActive();
    DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + ' backup ' + day, folder);
    props.setProperty('LAST_BACKUP', day);
    const files = []; const fi = folder.getFiles(); while (fi.hasNext()) files.push(fi.next());
    files.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(7).forEach(f => f.setTrashed(true));
  } catch (x) { console.error('Backup failed: ' + x.message); }
}

function saveAll(d, baseRev, deviceId, deviceName) {
  if (!d || !Array.isArray(d.txns)) throw fail('Invalid data');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rev = getRev();
    if (baseRev !== null && baseRev !== undefined && Number(baseRev) !== rev)
      throw fail('Your data was changed on another device or in the Sheet. Reload to get the latest.', 'CONFLICT');

    const oldIds = {};
    readTab('Transactions', Session.getScriptTimeZone()).forEach(t => { oldIds[t.id] = 1; });
    const oldN = Object.keys(oldIds).length;
    // Safety net: never let a save wipe out most of the existing data (e.g. built from a partial read)
    if (oldN >= 5 && d.txns.length < oldN * 0.5)
      throw fail('Refused to save: it would delete ' + (oldN - d.txns.length) + ' of ' + oldN + ' transactions. Reload to get the latest data.', 'SHRINK');
    backupSheet();
    const accName = {};
    (d.accounts || []).forEach(a => { accName[a.id] = a.name; });
    const snaps = [];
    (d.snapshots || []).slice().sort((a, b) => a.month < b.month ? -1 : 1).forEach(s =>
      Object.keys(s.values || {}).forEach(id => snaps.push({ month: s.month, accountId: id, account: accName[id] || '', value: s.values[id], updated: (s.updated || {})[id] || '' })));
    const cats = [];
    ['expense', 'income'].forEach(type => ((d.categories || {})[type] || []).forEach(c => cats.push({ type: type, name: c.name, icon: c.icon, color: c.color })));
    const settings = Object.keys(d.settings || {}).map(k => ({ key: k, value: JSON.stringify(d.settings[k]) }));

    writeTab('Transactions', d.txns.slice().sort((a, b) => a.date < b.date ? 1 : a.date > b.date ? -1 : (b.ts || 0) - (a.ts || 0)));
    writeTab('Accounts', d.accounts || []);
    writeTab('Snapshots', snaps);
    writeTab('EMIs', d.emis || []);
    writeTab('Recurring', d.recurring || []);
    writeTab('Categories', cats);
    writeTab('Settings', settings);
    const newRev = bumpRev();
    try { notifyNewTxns(d.txns.filter(t => t.id && !oldIds[t.id] && t.type !== 'transfer'), deviceId, deviceName); } catch (x) { console.error('Push failed: ' + x.message); }
    return { ok: true, ver: VERSION, rev: newRev };
  } finally {
    lock.releaseLock();
  }
}

/* ======================= Push notifications (Firebase Cloud Messaging) ======================= */
function getDevices() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('DEVICES') || '{}'); } catch (x) { return {}; } }
function putDevices(d) { PropertiesService.getScriptProperties().setProperty('DEVICES', JSON.stringify(d)); }

function registerDevice(req) {
  if (!req.deviceId || !req.token) throw fail('Missing device details');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = getDevices();
    Object.keys(d).forEach(k => { if (d[k].token === req.token && k !== req.deviceId) delete d[k]; });
    d[req.deviceId] = { token: req.token, name: String(req.deviceName || '').slice(0, 30), ts: Date.now() };
    putDevices(d);
  } finally { lock.releaseLock(); }
  return { ok: true, configured: !!fcmKey() };
}

function unregisterDevice(id) {
  const d = getDevices(); delete d[id]; putDevices(d); return { ok: true };
}

function fcmKey() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('FCM_SA') || ''); } catch (x) { return null; } }

function setFcmKeyPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Firebase service account', 'Paste the full contents of the service-account JSON key file downloaded from Firebase (see SETUP.md, push notifications).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  let k; try { k = JSON.parse(r.getResponseText()); } catch (x) { ui.alert('That is not valid JSON.'); return; }
  if (!k.private_key || !k.client_email || !k.project_id) { ui.alert('This does not look like a service-account key (needs project_id, client_email, private_key).'); return; }
  PropertiesService.getScriptProperties().setProperty('FCM_SA', JSON.stringify(k));
  CacheService.getScriptCache().remove('fcm_token');
  ui.alert('Saved ✓ Project: ' + k.project_id + '\n\nUse Paisa → Send test notification to check it.');
}

function fcmAccessToken(k) {
  const cache = CacheService.getScriptCache(), hit = cache.get('fcm_token');
  if (hit) return hit;
  const b64 = o => Utilities.base64EncodeWebSafe(typeof o === 'string' ? o : JSON.stringify(o)).replace(/=+$/, '');
  const now = Math.floor(Date.now() / 1000);
  const head = b64({ alg: 'RS256', typ: 'JWT' });
  const claim = b64({ iss: k.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging', aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600 });
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(head + '.' + claim, k.private_key)).replace(/=+$/, '');
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: head + '.' + claim + '.' + sig } });
  const j = JSON.parse(r.getContentText());
  if (!j.access_token) throw new Error('Firebase auth failed: ' + (j.error_description || j.error));
  cache.put('fcm_token', j.access_token, 3000);
  return j.access_token;
}

// Sends one data-only message to every registered device except the sender. Returns number sent.
function pushToOthers(title, body, exceptId, url, tag) {
  const k = fcmKey(); if (!k) return 0;
  const devices = getDevices(), ids = Object.keys(devices).filter(id => id !== exceptId);
  if (!ids.length) return 0;
  const auth = fcmAccessToken(k);
  const reqs = ids.map(id => ({ url: 'https://fcm.googleapis.com/v1/projects/' + k.project_id + '/messages:send', method: 'post', muteHttpExceptions: true,
    contentType: 'application/json', headers: { Authorization: 'Bearer ' + auth },
    payload: JSON.stringify({ message: { token: devices[id].token, data: { title: title, body: body, tag: tag || 'paisa-txn', url: url || './' }, webpush: { headers: { Urgency: 'high', TTL: '86400' } } } }) }));
  const res = UrlFetchApp.fetchAll(reqs);
  let sent = 0, changed = false;
  res.forEach((r, i) => {
    const c = r.getResponseCode();
    if (c === 200) sent++;
    else if (c === 404 || c === 400 && /UNREGISTERED|INVALID_ARGUMENT/.test(r.getContentText())) { delete devices[ids[i]]; changed = true; }
  });
  if (changed) putDevices(devices);
  return sent;
}

function inr(n) { return '₹' + Number(n || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 }); }

function notifyNewTxns(list, deviceId, deviceName) {
  if (!list.length) return;
  const who = deviceName || 'Someone';
  let title, body;
  if (list.length === 1) {
    const t = list[0], inc = t.type === 'income';
    title = who + (inc ? ' added income ' : ' added an expense ') + inr(t.amount);
    body = [t.category, t.note].filter(Boolean).join(' · ') + (t.date ? ' · ' + t.date : '');
  } else {
    title = who + ' added ' + list.length + ' transactions';
    body = list.slice(0, 3).map(t => (t.type === 'income' ? '+' : '−') + inr(t.amount) + ' ' + (t.category || '')).join(', ') + (list.length > 3 ? '…' : '');
  }
  pushToOthers(title, body, deviceId);
}

function sendTestPush() {
  const ui = SpreadsheetApp.getUi();
  if (!fcmKey()) { ui.alert('Set the Firebase key first: Paisa → Set up push notifications.'); return; }
  const n = Object.keys(getDevices()).length;
  if (!n) { ui.alert('No devices registered yet. Turn on Notifications in the app (Settings) first.'); return; }
  const sent = pushToOthers('Paisa test 🔔', 'Notifications are working.', '');
  ui.alert('Sent to ' + sent + ' of ' + n + ' device(s).');
}

function writeTab(name, rows) {
  const cols = SCHEMA[name], sh = tab(name);
  const data = [cols.map(c => c[0])].concat(rows.map(r => cols.map(([k, t]) => toCell(r[k], t))));
  if (sh.getLastRow() === data.length && sh.getLastColumn() >= cols.length) { // unchanged tab: skip the slow rewrite
    const cur = sh.getRange(1, 1, data.length, cols.length).getValues();
    if (cur.every((r, i) => r.every((v, j) => String(v) === String(data[i][j])))) return;
  }
  const n = Math.max(data.length, 2);
  // Text columns stay text (so dates/months are not auto-converted); amounts stay numeric for your own formulas.
  cols.forEach(([, t], i) => sh.getRange(1, i + 1, n, 1).setNumberFormat(t === 'n' ? '#,##0.00' : t === 'i' ? '0' : '@'));
  sh.getRange(1, 1, data.length, cols.length).setValues(data);
  const last = sh.getLastRow(); // clear only leftover old rows, after the new data is in place
  if (last > data.length) sh.getRange(data.length + 1, 1, last - data.length, Math.max(sh.getLastColumn(), cols.length)).clearContent();
  if (sh.getMaxRows() > data.length + 50) sh.deleteRows(data.length + 51, sh.getMaxRows() - data.length - 50);
  styleHeader(sh, cols.length);
}

function toCell(v, t) {
  if (v === undefined || v === null) return t === 'n' || t === 'i' ? 0 : t === 'b' ? 'TRUE' : '';
  if (t === 'a') return (Array.isArray(v) ? v : [v]).join(', ');
  if (t === 'b') return v === false ? 'FALSE' : 'TRUE';
  if (t === 'n' || t === 'i') return Number(v) || 0;
  return String(v);
}

/* ======================= Helpers ======================= */
function tab(name) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(name) || ss.insertSheet(name); }

function styleHeader(sh, n) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#ede9fe');
}

function getRev() { return Number(PropertiesService.getScriptProperties().getProperty('REV') || 0); }
function bumpRev() { const r = getRev() + 1; PropertiesService.getScriptProperties().setProperty('REV', String(r)); return r; }

function fail(msg, code) { const e = new Error(msg); e.code = code || ''; return e; }

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }


/* ======================= Reminders (hourly trigger) ======================= */
// Settings live in the app (Settings → Reminders) and arrive via the Settings tab as the "reminders" key.
function installReminders() {
  removeReminders(true);
  ScriptApp.newTrigger('reminderTick').timeBased().everyHours(1).create();
  const ui = SpreadsheetApp.getUi();
  ui.alert('Reminders enabled ✓\n\nChecked every hour. Times are in the script time zone: ' + Session.getScriptTimeZone() + '.\nIf that is not your time zone, set it in Project Settings → Time zone.');
}

function removeReminders(quiet) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'reminderTick').forEach(t => ScriptApp.deleteTrigger(t));
  if (!quiet) SpreadsheetApp.getUi().alert('Reminders disabled.');
}

function getReminderSettings() {
  const def = { due: 2, dueHour: 9, daily: true, dailyHour: 21 };
  const row = readTab('Settings', Session.getScriptTimeZone()).filter(r => r.key === 'reminders')[0];
  const v = row ? parseVal(row.value) : null;
  return Object.assign(def, v && typeof v === 'object' ? v : {});
}

function reminderTick() {
  if (!fcmKey() || !Object.keys(getDevices()).length) return;
  const tz = Session.getScriptTimeZone(), now = new Date(), hour = Number(Utilities.formatDate(now, tz, 'H')), day = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const cfg = getReminderSettings(), props = PropertiesService.getScriptProperties();
  if (cfg.due > 0 && hour >= cfg.dueHour && props.getProperty('LAST_DUE') !== day) {
    props.setProperty('LAST_DUE', day);
    sendDueReminder(Math.max(0, cfg.due - 1), tz, now);
  }
  if (cfg.daily && hour >= cfg.dailyHour && props.getProperty('LAST_DAILY') !== day) {
    props.setProperty('LAST_DAILY', day);
    const todays = readTab('Transactions', tz).some(t => t.date === day && t.type !== 'transfer');
    if (!todays) pushToOthers('📝 Add today\'s expenses', 'Nothing logged today yet. Tap to add.', '', './?add=expense', 'paisa-daily');
  }
}

// Items due today, within `ahead` days, or overdue this month. Skips anything already recorded this month.
function dueItems(ahead, tz, now) {
  const ym = Utilities.formatDate(now, tz, 'yyyy-MM'), today = Number(Utilities.formatDate(now, tz, 'd'));
  const dim = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const txns = readTab('Transactions', tz), out = [];
  const diff = dd => Math.min(Number(dd) || 1, dim) - today;
  const label = d => d < 0 ? 'overdue ' + (-d) + 'd' : d === 0 ? 'today' : d === 1 ? 'tomorrow' : 'in ' + d + 'd';
  const mdiff = (a, b) => (Number(b.slice(0, 4)) - Number(a.slice(0, 4))) * 12 + Number(b.slice(5, 7)) - Number(a.slice(5, 7));
  readTab('EMIs', tz).filter(e => e.id && e.name).forEach(e => {
    const i = mdiff(e.start, ym), paid = (Number(e.prepaid) || 0) + txns.filter(t => t.emiId === e.id).length;
    if (i < 0 || i >= e.tenure || paid > i) return;
    const d = paid < Math.min(i, e.tenure) ? -1 : diff(e.dueDay);
    if (d <= ahead) out.push({ d: d, text: 'EMI ' + e.name + ' ' + inr(e.amount) + ' (' + label(d) + ')' });
  });
  readTab('Recurring', tz).filter(r => r.id && r.active && r.type !== 'income' && (!r.since || r.since <= ym)).forEach(r => {
    if (txns.some(t => t.recurId === r.id && String(t.date).slice(0, 7) === ym)) return;
    const d = diff(r.dueDay);
    if (d <= ahead) out.push({ d: d, text: r.name + ' ' + inr(r.amount) + ' (' + label(d) + ')' });
  });
  readTab('Accounts', tz).filter(a => a.group === 'Credit Card' && a.dueDay).forEach(a => {
    if (txns.some(t => t.type === 'transfer' && t.toId === a.id && String(t.date).slice(0, 7) === ym)) return;
    const d = diff(a.dueDay);
    if (d >= 0 && d <= ahead) out.push({ d: d, text: a.name + ' card bill (' + label(d) + ')' });
  });
  return out.sort((x, y) => x.d - y.d);
}

function sendDueReminder(ahead, tz, now) {
  const items = dueItems(ahead, tz, now);
  if (!items.length) return 0;
  const title = items.length === 1 ? '📅 1 payment due' : '📅 ' + items.length + ' payments due';
  const body = items.slice(0, 4).map(i => i.text).join('\n') + (items.length > 4 ? '\n+' + (items.length - 4) + ' more' : '');
  return pushToOthers(title, body, '', './', 'paisa-due');
}


/* ======================= Bank email alerts → "To review" inbox ======================= */
// Every 30 min, Gmail is searched for bank alert emails. Parsed transactions land in the Inbox tab
// (not in Transactions): the app shows them under "To review" and you add or skip each one.
const BANK_QUERY_DEFAULT = 'newer_than:3d -in:spam -in:trash -category:promotions -category:social {debited credited spent withdrawn "transaction alert" "txn alert" "has been used"}';
const BANKS = [['hdfc', 'HDFC'], ['icici', 'ICICI'], ['sbi', 'SBI'], ['axis', 'Axis'], ['kotak', 'Kotak'], ['idfc', 'IDFC First'], ['yesbank', 'Yes Bank'],
  ['indusind', 'IndusInd'], ['aubank', 'AU'], ['federal', 'Federal'], ['rbl', 'RBL'], ['hsbc', 'HSBC'], ['sc.com', 'Standard Chartered'], ['standardchartered', 'Standard Chartered'],
  ['americanexpress', 'Amex'], ['aexp', 'Amex'], ['citi', 'Citi'], ['pnb', 'PNB'], ['bankofbaroda', 'Bank of Baroda'], ['bobcard', 'BOB Card'], ['bobfinancial', 'BOB Card'], ['scapia', 'Scapia'], ['cred.club', 'CRED'], ['swiggy', 'Swiggy'], ['flipkart', 'Flipkart'], ['canara', 'Canara'],
  ['unionbank', 'Union Bank'], ['idbi', 'IDBI'], ['indianbank', 'Indian Bank'], ['dbs', 'DBS'], ['onecard', 'OneCard'], ['slice', 'Slice'], ['jupiter', 'Jupiter'], ['fi.money', 'Fi'], ['paytm', 'Paytm']];

// Sender addresses you add via the menu: every email from them is checked (any Gmail tab), no keywords needed.
function bankSenders() { return (PropertiesService.getScriptProperties().getProperty('BANK_SENDERS') || '').split(',').map(x => x.trim()).filter(Boolean); }
function senderQuery() {
  const a = bankSenders(); if (!a.length) return '';
  return 'newer_than:3d -in:spam -in:trash from:(' + a.join(' OR ') + ')';
}

function setBankSendersPrompt() {
  const ui = SpreadsheetApp.getUi(), cur = bankSenders();
  const r = ui.prompt('Bank alert sender emails', 'Enter the email addresses your bank alerts come from, separated by commas.\nEmails from these senders are always checked, in any Gmail tab.\n\nExample:\nalerts@axis.bank.in, alerts@hdfcbank.net, credit_cards@icicibank.com\n\nCurrent: ' + (cur.join(', ') || '(none)') + '\nLeave empty to clear.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const list = r.getResponseText().split(/[,;\s]+/).map(x => x.trim().toLowerCase()).filter(x => /^[^@\s]+@[^@\s]+$|^@?[\w.-]+\.[a-z]{2,}$/.test(x)).map(x => x.replace(/^@/, ''));
  const props = PropertiesService.getScriptProperties();
  if (list.length) props.setProperty('BANK_SENDERS', list.join(',')); else props.deleteProperty('BANK_SENDERS');
  ui.alert(list.length ? 'Saved ✓\n\nWatching: ' + list.join(', ') + '\n\nRun "Import bank emails now" to test.' : 'Cleared. Using the default keyword search only.');
}

function bankQuery() { return PropertiesService.getScriptProperties().getProperty('BANK_QUERY') || BANK_QUERY_DEFAULT; }

function installBankImport() {
  removeBankImport(true);
  ScriptApp.newTrigger('bankTick').timeBased().everyMinutes(30).create();
  ensureStatementTrigger();
  const n = importBankEmails();
  try { importStatements(); } catch (e) { /* Drive permission is asked on first statement import */ }
  SpreadsheetApp.getUi().alert('Bank email import enabled ✓\n\nGmail is checked every 30 minutes for bank alert emails.\nFound ' + n + ' new transaction(s) from the last 3 days — open the app to review them.\n\nTip: turn on email alerts for every transaction in your bank / card app.');
}

function removeBankImport(quiet) {
  ScriptApp.getProjectTriggers().filter(t => ['bankTick', 'statementTick'].includes(t.getHandlerFunction())).forEach(t => ScriptApp.deleteTrigger(t));
  if (!quiet) SpreadsheetApp.getUi().alert('Bank email import disabled.');
}

function importBankEmailsNow() {
  const diag = [], n = importBankEmails(diag);
  SpreadsheetApp.getUi().alert('Added ' + n + ' new bank transaction(s) to review.\n\nEmails checked (' + diag.length + '):\n' +
    (diag.length ? diag.slice(0, 20).join('\n') : '— none matched the Gmail search —'));
}

function setBankQueryPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Bank email search', 'Gmail search used to find alert emails. Leave empty to reset to the default.\n\nCurrent:\n' + bankQuery() +
    '\n\nExample (only some senders):\nnewer_than:3d {from:alerts@hdfcbank.net from:credit_cards@icicibank.com}', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const q = r.getResponseText().trim(), props = PropertiesService.getScriptProperties();
  if (q) props.setProperty('BANK_QUERY', q); else props.deleteProperty('BANK_QUERY');
  ui.alert('Saved ✓');
}

function ensureStatementTrigger() {
  if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'statementTick')) return;
  ScriptApp.newTrigger('statementTick').timeBased().everyDays(1).atHour(7).create();
}

// Card statements are scanned once a day (heavier: Gmail PDFs + Drive); bank alerts stay on the 30-minute bankTick.
function statementTick() {
  let k = 0; try { k = importStatements(); } catch (e) { console.error('Statements: ' + e.message); }
  if (k) pushToOthers('📄 ' + k + ' card statement' + (k === 1 ? '' : 's') + ' saved', 'Open Paisa → Credit Cards.', '', './', 'paisa-stmt');
}

function bankTick() {
  const n = importBankEmails();
  if (n) pushToOthers('🏦 ' + n + ' bank transaction' + (n === 1 ? '' : 's') + ' to review', 'Tap to add them to Paisa.', '', './?review=1', 'paisa-bank');
  ensureStatementTrigger(); // upgrades older installs: statements now run once a day
}

/* ======================= Credit card statements → Drive folder + "Statements" tab ======================= */
// Credit card statement emails with a PDF attachment are copied to the Drive folder "Paisa Statements".
// Each one is listed in the Statements tab; the app shows the list and opens the file from Drive on tap.
// Card issuers' sender domains (all major Indian banks + co-brand partners). Also matched by statement wording below.
const STMT_SENDERS = ['axisbank.com', 'axis.bank.in', 'hdfcbank.net', 'hdfcbank.com', 'icicibank.com', 'sbicard.com', 'sbi.co.in', 'hsbc.co.in', 'hsbc.com', 'indusind.com', 'indusindbank.com',
  'kotak.com', 'idfcfirstbank.com', 'yesbank.in', 'aubank.in', 'rblbank.com', 'federalbank.co.in', 'scapia.cc', 'scapia.in', 'bobcard.co.in', 'bobfinancial.com', 'bankofbaroda.com', 'bankofbaroda.co.in',
  'sc.com', 'americanexpress.com', 'aexp.com', 'citi.com', 'citibank.com', 'pnb.co.in', 'canarabank.com', 'unionbankofindia.co.in', 'idbi.co.in', 'onecard.in', 'getonecard.app', 'sliceit.com', 'cred.club', 'flipkart.com', 'swiggy.in'];
const STMT_WORDS = '{statement "e-statement" estatement "card bill" "credit card bill" "billing statement"}';
const STMT_WINDOW = 'newer_than:120d';
// Product name detection → label shown in the app. First match wins.
const STMT_PRODUCTS = [[/flipkart[\s-]*neo|\bneo\b.*(?:card|axis)/i, 'Flipkart Neo'], [/flipkart/i, 'Axis Flipkart'], [/swiggy/i, 'Swiggy HDFC'], [/scapia/i, 'Scapia BOB'],
  [/hsbc/i, 'HSBC Platinum'], [/sbi\s*card|sbicard|cashback\s*sbi|sbi\s*cashback/i, 'SBI Cashback'], [/\bcred\b|indusind/i, 'Cred IndusInd'], [/amazon\s*pay|amazon/i, 'Amazon Pay ICICI']];
const STMT_FOLDER_NAME = 'Paisa Statements';

function stmtFolder() {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty('STMT_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* deleted: recreate */ } }
  const it = DriveApp.getFoldersByName(STMT_FOLDER_NAME), f = it.hasNext() ? it.next() : DriveApp.createFolder(STMT_FOLDER_NAME);
  props.setProperty('STMT_FOLDER', f.getId()); return f;
}

function stmtSheet() {
  const sh = tab('Statements'), cols = SCHEMA.Statements;
  if (sh.getLastRow() > 0 && sh.getLastColumn() < cols.length) sh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c[0])]);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c[0])]);
    cols.forEach(([, t], i) => sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat(t === 'n' ? '#,##0.00' : t === 'i' ? '0' : '@'));
    styleHeader(sh, cols.length);
  }
  return sh;
}

function stmtDate(str) {
  const m = String(str || '').match(/(\d{1,2})[-\/ ]([A-Za-z]{3,9}|\d{1,2})[-\/ ,]*(\d{2,4})/); if (!m) return '';
  let mo = /^\d+$/.test(m[2]) ? +m[2] : ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(m[2].slice(0, 3).toLowerCase()) + 1;
  const y = +m[3] < 100 ? 2000 + +m[3] : +m[3];
  return mo >= 1 && mo <= 12 ? y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
}

function parseStatementMail(subject, body, from, fileName) {
  const text = (String(subject || '') + ' . ' + String(body || '')).replace(/\s+/g, ' ').slice(0, 6000), hay = text + ' ' + fileName;
  const prod = (STMT_PRODUCTS.find(([r]) => r.test(hay + ' ' + from)) || [])[1] || '';
  const fromS0 = String(from || '').toLowerCase(), known = STMT_SENDERS.some(d => fromS0.indexOf(d) >= 0);
  if (!(/\bcard\b/i.test(hay) || prod) || !/statement|e-?stmt|card bill|billing/i.test(hay)) return null;
  if (!prod && !known && !/credit\s*card/i.test(hay)) return null;
  if (/savings|current account|demat|mutual fund|folio|loan|portfolio|account statement/i.test((subject || '') + ' ' + fileName) && !/card/i.test((subject || '') + ' ' + fileName)) return null;
  const amt = re => { const m = text.match(re); return m ? Number(m[1].replace(/,/g, '')) || 0 : 0; };
  const money = '[^0-9]{0,25}?(?:INR|Rs\\.?|₹)?\\s*([0-9][0-9,]*(?:\\.[0-9]{1,2})?)';
  const due = amt(new RegExp('total (?:amount |payment )?due' + money, 'i')) || amt(new RegExp('total amount payable' + money, 'i'));
  const min = amt(new RegExp('min(?:imum)? (?:amount |payment )?due' + money, 'i'));
  const dm = text.match(/(?:payment )?due (?:date|on)[^0-9A-Za-z]{0,6}(\d{1,2}[-\/ ][A-Za-z0-9]{2,9}[-\/ ,]*\d{2,4})/i);
  const acct = hay.match(/(?:card|ending|no\.?)[\s:]*[xX*•.\-]*(\d{4})\b/i);
  const addr = (String(from || '').match(/<([^>]+)>/) || [0, String(from || '')])[1].toLowerCase();
  const bank = (BANKS.find(([k]) => addr.indexOf(k) >= 0) || BANKS.find(([k]) => hay.toLowerCase().indexOf(k) >= 0 || (subject || '').toLowerCase().indexOf(k) >= 0) || [])[1] || (prod ? prod.split(' ')[0] : 'Bank');
  const pi = text.search(/\bpassword\b/i), pre = text.slice(Math.max(0, pi - 100), Math.max(pi, 0)), pk = Math.max(pre.lastIndexOf('. '), pre.lastIndexOf('! ')), pw = pi >= 0 ? [pre.slice(pk + 2 > 1 && pk >= 0 ? pk + 2 : 0) + text.slice(pi, pi + 260)] : null;
  return { bank: bank, card: acct ? acct[1] : '', amount: due, minDue: min, dueDate: dm ? stmtDate(dm[1]) : '', pwHint: pw ? pw[0].trim().slice(0, 260) : '', product: prod };
}

// Returns the number of new statement PDFs saved. diag (optional array) gets one line per email.
function importStatements(diag) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const tz = Session.getScriptTimeZone(), sh = stmtSheet(), seen = {}, thSeen = {}, rows = [], threads = [];
    readTab('Statements', tz).forEach(r => { seen[r.id] = 1; if (r.hash) seen['h' + r.hash] = 1; });
    const sq = senderQuery(), base = STMT_WINDOW + ' has:attachment filename:pdf -in:spam -in:trash ';
    const qs = [base + STMT_WORDS, base + 'from:(' + STMT_SENDERS.join(' OR ') + ')', base + '{"credit card" "card ending" "card no"}'];
    if (sq) qs.push(sq.replace('newer_than:3d', STMT_WINDOW + ' has:attachment filename:pdf'));
    qs.forEach(q => {
      try { GmailApp.search(q, 0, 60).forEach(th => { if (!thSeen[th.getId()]) { thSeen[th.getId()] = 1; threads.push(th); } }); }
      catch (e) { if (diag) diag.push('⚠️ Search failed: ' + e.message); }
    });
    let folder = null;
    threads.forEach(th => th.getMessages().forEach(m => {
      const subj = String(m.getSubject() || '(no subject)').slice(0, 60);
      m.getAttachments({ includeInlineImages: false, includeAttachments: true }).forEach((att, i) => {
        const orig = att.getName() || 'statement.pdf';
        if (!/\.pdf$/i.test(orig) && att.getContentType() !== 'application/pdf') return;
        const id = m.getId() + '_' + i;
        if (seen[id]) { if (diag) diag.push('• ' + subj + ' → already saved'); return; }
        const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, att.getBytes())).slice(0, 22);
        if (seen['h' + hash]) { if (diag) diag.push('• ' + subj + ' → duplicate of a saved PDF'); return; }
        const p = parseStatementMail(m.getSubject(), m.getPlainBody() || m.getBody().replace(/<[^>]+>/g, ' '), m.getFrom(), orig);
        if (!p) { if (diag) diag.push('• ' + subj + ' (' + orig + ') → ✗ not a credit card statement'); return; }
        seen[id] = 1; seen['h' + hash] = 1;
        const d = m.getDate(), ds = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        folder = folder || stmtFolder();
        const name = ((p.product || p.bank) + ' card' + (p.card ? ' ' + p.card : '') + ' ' + ds.slice(0, 7) + ' - ' + orig).replace(/[\\/:*?"<>|]/g, '-');
        const file = folder.createFile(att.copyBlob().setName(name));
        rows.push(Object.assign(p, { hash: hash, hidden: '', id: id, date: ds, name: name, fileId: file.getId(), url: file.getUrl(), ts: d.getTime() }));
        if (diag) diag.push('• ' + subj + ' → ✓ saved ' + name);
      });
    }));
    if (rows.length) {
      const cols = SCHEMA.Statements;
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, cols.length).setValues(rows.map(r => cols.map(([k, t]) => toCell(r[k], t))));
    }
    return rows.length;
  } finally { lock.releaseLock(); }
}

function importStatementsNow() {
  const diag = [], n = importStatements(diag);
  SpreadsheetApp.getUi().alert('Saved ' + n + ' new statement(s) to the Drive folder "' + STMT_FOLDER_NAME + '".\n\nEmails checked (' + diag.length + '):\n' +
    (diag.length ? diag.slice(0, 20).join('\n') : '— none matched the Gmail search —'));
}

function readStatements() {
  return readTab('Statements', Session.getScriptTimeZone()).filter(r => r.id && r.url && !r.hidden).sort((a, b) => b.ts - a.ts).slice(0, 80)
    .map(r => ({ id: r.id, date: r.date, bank: r.bank, card: r.card, amount: r.amount, minDue: r.minDue, dueDate: r.dueDate, name: r.name, url: r.url, pwHint: r.pwHint, product: r.product, ts: r.ts }));
}

// Returns the number of new items added to the Inbox tab.
// diag (optional array) collects one line per email explaining what happened.
function importBankEmails(diag) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const tz = Session.getScriptTimeZone(), sh = inboxSheet(), seen = {}, thSeen = {};
    readTab('Inbox', tz).forEach(r => { seen[r.id] = 1; });
    const rows = [], threads = [];
    [bankQuery(), senderQuery()].filter(Boolean).forEach(q => {
      try { GmailApp.search(q, 0, 50).forEach(th => { if (!thSeen[th.getId()]) { thSeen[th.getId()] = 1; threads.push(th); } }); }
      catch (e) { if (diag) diag.push('⚠️ Search failed: ' + e.message); }
    });
    threads.forEach(th => th.getMessages().forEach(m => {
      const id = m.getId(), subj = String(m.getSubject() || '(no subject)').slice(0, 60);
      if (seen[id]) { if (diag) diag.push('• ' + subj + ' → already imported'); return; }
      seen[id] = 1;
      const p = parseBankEmail(m.getSubject(), m.getPlainBody() || m.getBody().replace(/<[^>]+>/g, ' '), m.getFrom());
      if (diag) diag.push('• ' + subj + ' → ' + (p ? '✓ ₹' + p.amount + ' ' + p.type + (p.merchant ? ' · ' + p.merchant : '') : '✗ ' + PARSE_WHY));
      if (!p) return;
      const d = m.getDate();
      rows.push(Object.assign(p, { id: id, date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'), status: '', txId: '', ts: d.getTime() }));
    }));
    if (rows.length) {
      const cols = SCHEMA.Inbox;
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, cols.length).setValues(rows.map(r => cols.map(([k, t]) => toCell(r[k], t))));
    }
    pruneInbox(sh);
    return rows.length;
  } finally { lock.releaseLock(); }
}

function inboxSheet() {
  const sh = tab('Inbox'), cols = SCHEMA.Inbox;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols.map(c => c[0])]);
    cols.forEach(([, t], i) => sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat(t === 'n' ? '#,##0.00' : t === 'i' ? '0' : '@'));
    styleHeader(sh, cols.length);
  }
  return sh;
}

// Keep 120 days of handled rows (needed so the same email is never imported twice).
function pruneInbox(sh) {
  const n = sh.getLastRow(); if (n < 2) return;
  const cut = Date.now() - 120 * 864e5, ts = sh.getRange(2, SCHEMA.Inbox.length, n - 1, 1).getValues();
  for (let i = ts.length - 1; i >= 0; i--) if (Number(ts[i][0]) && Number(ts[i][0]) < cut) sh.deleteRow(i + 2);
}

function readInbox() {
  return readTab('Inbox', Session.getScriptTimeZone()).filter(r => r.id && !r.status && r.amount > 0)
    .sort((a, b) => b.ts - a.ts).slice(0, 100);
}

// marks: [{ id, status: 'added'|'skipped', txId }]
function inboxMark(marks) {
  if (!Array.isArray(marks) || !marks.length) return { ok: true };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = inboxSheet(), n = sh.getLastRow(); if (n < 2) return { ok: true };
    const ids = sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]));
    const sCol = SCHEMA.Inbox.findIndex(c => c[0] === 'status') + 1;
    marks.forEach(m => {
      const i = ids.indexOf(String(m.id)); if (i < 0) return;
      sh.getRange(i + 2, sCol, 1, 2).setValues([[m.status === 'added' ? 'added' : 'skipped', String(m.txId || '')]]);
    });
    return { ok: true };
  } finally { lock.releaseLock(); }
}

// Pulls amount / debit-credit / merchant / account out of an Indian bank alert. Returns null if it isn't one (reason in PARSE_WHY).
var PARSE_WHY = '';
function parseBankEmail(subject, body, from) {
  PARSE_WHY = '';
  const text = (String(subject || '') + ' . ' + String(body || '')).replace(/\s+/g, ' ').slice(0, 4000);
  // Safety footers ("Never share your OTP…", "If this transaction was not initiated by you…") must not trigger rejection.
  const core = text.replace(/[^.!?]*\b(never|do not|don't|dont|not initiated|not done by you|not you|beware|fraud|report|block|disclose|share)\b[^.!?]*[.!?]?/gi, ' ');
  const rej = core.slice(0, 700).match(/\b(otp|one[- ]time password|verification code|statement (is )?(ready|generated)|payment (is )?due|due date|minimum (amount )?due|will be (debited|charged)|collect request|request(ed)? (money|payment)|(transaction|payment|txn) (has )?(failed|declined)|unsuccessful|e-?mandate (registered|created))\b/i);
  if (rej) { PARSE_WHY = 'skipped (mentions "' + rej[0] + '")'; return null; }

  let amount = 0;
  const amtRe = /(?:INR|Rs\.?|₹)\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/gi; let m;
  while ((m = amtRe.exec(text))) {
    const before = text.slice(Math.max(0, m.index - 30), m.index);
    if (/(bal|balance|limit|avl|available|outstanding)[^.]*$/i.test(before)) continue;
    amount = Number(m[1].replace(/,/g, '')); if (amount > 0) break;
  }
  if (!(amount > 0)) { PARSE_WHY = 'no ₹ amount found'; return null; }

  const deb = text.search(/\b(debited|spent|withdrawn|paid|purchase|deducted|charged|sent|used|using|debit of|txn of)\b/i);
  const cre = text.search(/\b(credited|received|deposited|refund(ed)?|reversed|reversal|cashback)\b/i);
  if (deb < 0 && cre < 0) { PARSE_WHY = 'no debit/credit word'; return null; }
  const type = cre >= 0 && (deb < 0 || cre < deb) ? 'income' : 'expense';

  const acctM = text.match(/(?:a\/c|acct|account|card)(?:\s*(?:no\.?|number|ending(?:\s+(?:in|with))?))?[\s:]*[xX*•.\-]*(\d{4})\b/i);
  const fromS = String(from || '').toLowerCase();
  const bank = (BANKS.find(([k]) => fromS.indexOf(k) >= 0) || [])[1] || String(from || '').replace(/<.*>/, '').replace(/["']/g, '').trim().slice(0, 30);
  const method = /credit card|card ending|card no|card xx|card \*\*|\bcard\b.*\bused\b/i.test(text) && !/debit card/i.test(text) ? 'Card'
    : /\bUPI\b|\bVPA\b|@ok|@ybl|@paytm|@upi|@ibl|@axl/i.test(text) ? 'UPI'
    : /\b(NEFT|IMPS|RTGS|net ?banking)\b/i.test(text) ? 'Bank Transfer' : /debit card|\bATM\b/i.test(text) ? (/\bATM\b/i.test(text) ? 'Cash' : 'Card') : 'Bank Transfer';

  let merchant = '';
  const vpa = text.match(/\bVPA\s+([\w.\-]+@[\w.\-]+)(?:\s+\(?([A-Za-z][A-Za-z0-9 &.'\-]{2,40}?)\)?(?=\s+on\b|\s*[.,]|\s+UPI|\s+Ref|$))?/i);
  if (vpa) merchant = vpa[2] && !/^on$/i.test(vpa[2]) ? vpa[2] : vpa[1].split('@')[0];
  const bad = v => /^\d|^(a|an|the|your|a\/c|ac|acct|account|card|xx|\*|rs|inr|bank|you|us|our|this|upi|imps|neft|rtgs|date|behalf|transaction|txn|purchase|payment)\b/i.test(v) || /\b(debit|credit) card\b|\bcard no\b/i.test(v);
  const pick = words => {
    const re = new RegExp('\\b(?:' + words + ')\\s+([A-Za-z0-9][A-Za-z0-9 &.\'*\\/\\-]{2,40}?)(?=\\s+(?:on|at|via|ref|for|using|with|from|dated|is|was|has|in|your|thru|through|avl|available|info|upi|towards)\\b|\\s*[.,(:;]|$)', 'gi');
    let mm; while ((mm = re.exec(text))) { const v = mm[1].trim(); if (!bad(v)) return v; }
    return '';
  };
  if (!merchant) { const ref = text.match(/\b(?:UPI|IMPS|NEFT|RTGS)\/[A-Z0-9]{2,6}\/\d{6,}\/([A-Za-z][A-Za-z0-9 .&'\-]{1,40}?)(?=\s+(?:if|on|ref|avl|available|the|for)\b|\s*[.,\/]|$)/i); if (ref) merchant = ref[1]; }
  if (!merchant) merchant = pick('at');
  if (!merchant) { const info = text.match(/\binfo[:\s\-]+([A-Za-z0-9\/ .*&\-]{3,40}?)(?=\.\s|\.?$|\s+(?:the|avl|available|on|ref)\b)/i); if (info) merchant = info[1].split('/').filter(x => /[A-Za-z]{3}/.test(x)).pop() || ''; }
  if (!merchant) merchant = pick(type === 'income' ? 'by|from' : 'to|towards');
  if (!merchant) merchant = pick('on|for');
  merchant = merchant.replace(/\s+/g, ' ').replace(/[.\-*\s]+$/, '').trim().slice(0, 40);

  return { type: type, amount: amount, merchant: merchant, bank: bank, acct: acctM ? acctM[1] : '', method: method, snippet: String(subject || '').slice(0, 120) };
}


// Edit one statement row from the app. Only the editable fields are accepted.
function stmtSave(x) {
  if (!x || !x.id) throw fail('Missing statement id');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = stmtSheet(), n = sh.getLastRow(); if (n < 2) return { ok: true };
    const cols = SCHEMA.Statements, ids = sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0])), i = ids.indexOf(String(x.id));
    if (i < 0) throw fail('Statement not found');
    ['bank', 'product', 'card', 'amount', 'minDue', 'dueDate', 'pwHint', 'hidden'].forEach(k => {
      if (x[k] === undefined) return;
      const ci = cols.findIndex(c => c[0] === k); if (ci < 0) return;
      sh.getRange(i + 2, ci + 1).setValue(toCell(k === 'hidden' ? (x[k] ? '1' : '') : x[k], cols[ci][1]));
    });
    return { ok: true, statements: readStatements() };
  } finally { lock.releaseLock(); }
}
