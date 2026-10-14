/**
 * Time Register — Apps Script API (Google Sheets as database).
 * Bound to the spreadsheet: Extensions > Apps Script. Run setup() once, type the admin's password in the Employees sheet,
 * then deploy as a web app (Execute as: Me, Access: Anyone). The page lives on GitHub Pages.
 * Every action except login needs a signed session token.
 */
const TZ = 'Europe/Copenhagen';
const SH = { LOG: 'Attendance Logs', EMP: 'Employees', AUD: 'Audit Log' };
const LOG_HEAD = ['Record ID', 'Employee Name', 'Event Date', 'Event Time', 'Event Type', 'Entry Source', 'Created At', 'Employee Email', 'Instant (UTC)', 'Status'];
const EMP_HEAD = ['Employee ID', 'Employee Name', 'Login Identifier', 'Role', 'Active', 'Password'];
const AUD_HEAD = ['Audit ID', 'Record ID', 'Employee Affected', 'Action', 'Previous Value', 'Updated Value', 'Changed By', 'Changed At', 'Reason'];

/* ---------- Setup ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from inside the spreadsheet (Extensions > Apps Script).');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  ensureSheet_(ss, SH.LOG, LOG_HEAD);
  ensureSheet_(ss, SH.EMP, EMP_HEAD);
  ensureSheet_(ss, SH.AUD, AUD_HEAD);
  const emp = ss.getSheetByName(SH.EMP);
  if (emp.getMaxColumns() >= 7) emp.getRange(1, 7, emp.getMaxRows(), 1).clearContent(); // old PIN Salt column
  if (emp.getLastRow() < 2) {
    emp.getRange(2, 1, 1, 6).setValues([['E001', 'Admin', 'admin', 'Administrator', 'Yes', '']]);
  }
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
}

function ensureSheet_(ss, name, head) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // Plain text, so Sheets never turns 011026 into 11026 or 08:45 into a time value.
  sh.getRange(2, 1, sh.getMaxRows() - 1, head.length).setNumberFormat('@');
}

/* ---------- Web endpoint ---------- */

const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() { return json_({ ok: true, data: 'Time Register API is running.' }); }

function doPost(e) {
  try {
    const p = JSON.parse((e.postData && e.postData.contents) || '{}');
    const handler = ACTIONS_[p.action];
    if (!handler) throw new Error('Unknown action.');
    if (p.action !== 'login') {
      const u = userKey_(userFromToken_(p.token));
      const row = employeeRows_().find(x => x.username === u);
      if (!row || !row.active) throw new Error('SESSION');
      CTX_ME = publicUser_(row);
    }
    return json_({ ok: true, data: handler(p) });
  } catch (err) {
    if (err.message === 'SESSION') return json_({ ok: false, code: 'SESSION', error: 'Your session has expired. Sign in again.' });
    return json_({ ok: false, error: err.message });
  }
}

const ACTIONS_ = {
  login: p => api_login(p),
  bootstrap: () => api_bootstrap(),
  status: () => api_status(),
  punch: p => api_punch(p.type),
  dayEvents: p => api_dayEvents(p.email, p.ymd),
  saveManual: p => api_saveManual(p),
  updateEvent: p => api_updateEvent(p),
  deleteEvent: p => api_deleteEvent(p),
  report: p => api_report(p),
  changePin: p => api_changePin(p),
  listEmployees: () => api_listEmployees(),
  addEmployee: p => api_addEmployee(p),
  setPin: p => api_setPin(p),
  setActive: p => api_setActive(p)
};

/* ---------- Helpers ---------- */

function sheet_(name) {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Setup has not been run. Run setup() once from the Apps Script editor.');
  const sh = SpreadsheetApp.openById(id).getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

function values_(name, cols) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, 1, n, cols).getValues();
}

function appendRows_(name, rows) {
  const sh = sheet_(name);
  const start = sh.getLastRow() + 1;
  const last = start + rows.length - 1;
  if (last > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), last - sh.getMaxRows() + 200);
  const rg = sh.getRange(start, 1, rows.length, rows[0].length);
  rg.setNumberFormat('@');
  rg.setValues(rows);
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

const fmtDate_ = d => Utilities.formatDate(d, TZ, 'ddMMyy');
const fmtTime_ = d => Utilities.formatDate(d, TZ, 'HH:mm');
const ymd_ = t => Utilities.formatDate(new Date(t), TZ, 'yyyy-MM-dd');
const dmy_ = ymd => ymd.slice(8, 10) + '/' + ymd.slice(5, 7) + '/' + ymd.slice(0, 4);
const when_ = e => Utilities.formatDate(new Date(e.t), TZ, 'dd/MM HH:mm');
const cmp_ = (a, b) => a.t - b.t || (a.type === b.type ? 0 : (a.type === 'Check In' ? -1 : 1));

function fmtDur_(ms) {
  const m = Math.round(ms / 60000);
  return Math.floor(m / 60) + 'h ' + ('0' + (m % 60)).slice(-2) + 'm';
}

/** Local Copenhagen date + time -> Date. Handles daylight saving; rejects times that do not exist. */
function parseLocal_(ymd, hm) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd || '') || !/^\d{2}:\d{2}$/.test(hm || '')) throw new Error('Enter a valid date and time.');
  const d = Utilities.parseDate(ymd + ' ' + hm, TZ, 'yyyy-MM-dd HH:mm');
  if (Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm') !== ymd + ' ' + hm) {
    throw new Error(hm + ' on ' + dmy_(ymd) + ' does not exist (clocks change that night). Pick another time.');
  }
  return d;
}

function noFuture_(d) {
  if (d.getTime() > Date.now() + 60000) throw new Error('You cannot register a time in the future.');
}

function need_(reason) {
  const r = String(reason || '').trim();
  if (!r) throw new Error('Enter a reason for the correction.');
  return r;
}

/* ---------- Identity: username + PIN, signed session token ---------- */

const SESSION_DAYS = 7, MAX_FAILS = 5, LOCK_SECONDS = 900;
let CTX_ME = null;

const userKey_ = u => String(u || '').trim().toLowerCase();

function checkPinFormat_(pin) {
  const v = String(pin || '').trim();
  if (v.length < 4 || v.length > 40) throw new Error('A password must be 4 to 40 characters.');
}

function sign_(s) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) throw new Error('Setup has not been run. Run setup() once from the Apps Script editor.');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret));
}

function makeToken_(username) {
  const b = Utilities.base64EncodeWebSafe(username + '|' + (Date.now() + SESSION_DAYS * 86400000));
  return b + '.' + sign_(b);
}

function userFromToken_(tok) {
  const parts = String(tok || '').split('.');
  if (parts.length !== 2 || sign_(parts[0]) !== parts[1]) throw new Error('SESSION');
  const bits = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString().split('|');
  if (Date.now() > Number(bits[1])) throw new Error('SESSION');
  return bits[0];
}

function employeeRows_() {
  const sh = sheet_(SH.EMP);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 6).getValues().map((r, i) => ({
    row: i + 2, id: r[0], name: r[1], username: userKey_(r[2]), role: String(r[3]).trim(),
    active: String(r[4]).trim().toLowerCase() === 'yes', pin: String(r[5]).trim()
  }));
}

const publicUser_ = e => ({ id: e.id, name: e.name, email: e.username, admin: e.role.toLowerCase() === 'administrator' });

function me_() {
  if (!CTX_ME) throw new Error('SESSION');
  return CTX_ME;
}

function employees_() {
  return employeeRows_().filter(e => e.active).map(e => ({ id: e.id, name: e.name, email: e.username }));
}

function api_login(p) {
  const u = userKey_(p.username);
  const cache = CacheService.getScriptCache();
  const key = 'fail:' + u;
  if (Number(cache.get(key) || 0) >= MAX_FAILS) throw new Error('Too many wrong attempts. Try again in 15 minutes.');
  const e = employeeRows_().find(x => x.username === u);
  const ok = e && e.active && e.pin && String(p.pin || '').trim() === e.pin;
  if (!ok) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), LOCK_SECONDS);
    throw new Error('Wrong username or password.');
  }
  cache.remove(key);
  return { token: makeToken_(u), user: publicUser_(e) };
}

function requireAdmin_() {
  const me = me_();
  if (!me.admin) throw new Error('Only administrators can do this.');
  return me;
}

function writePin_(row, pin) {
  const cell = sheet_(SH.EMP).getRange(row, 6);
  cell.setNumberFormat('@');
  cell.setValue(String(pin).trim());
}

function addEmployee_(name, username, role, pin, allowExisting) {
  const u = userKey_(username);
  if (!String(name || '').trim()) throw new Error('Enter a name.');
  if (!/^[a-z0-9._@+-]{2,60}$/.test(u)) throw new Error('Username: 2 to 60 characters, using letters, digits and . _ @ + -');
  checkPinFormat_(pin);
  const r = String(role).toLowerCase() === 'administrator' ? 'Administrator' : 'Employee';
  const ex = employeeRows_().find(x => x.username === u);
  if (ex) {
    if (!allowExisting) throw new Error('That username is already in use.');
    sheet_(SH.EMP).getRange(ex.row, 2, 1, 5).setValues([[String(name).trim(), u, r, 'Yes', String(pin).trim()]]);
    return;
  }
  const n = employeeRows_().length + 1;
  appendRows_(SH.EMP, [['E' + ('00' + n).slice(-3), String(name).trim(), u, r, 'Yes', String(pin).trim()]]);
}

function api_changePin(p) {
  const me = me_();
  checkPinFormat_(p.newPin);
  return locked_(() => {
    const e = employeeRows_().find(x => x.username === me.email);
    if (!e || String(p.oldPin || '').trim() !== e.pin) throw new Error('Your current password is wrong.');
    writePin_(e.row, p.newPin);
    return 'Password changed.';
  });
}

function api_listEmployees() {
  requireAdmin_();
  return employeeRows_().map(e => ({ id: e.id, name: e.name, email: e.username, role: e.role, active: e.active, hasPin: !!e.pin }));
}

function api_addEmployee(p) {
  requireAdmin_();
  return locked_(() => { addEmployee_(p.name, p.email, p.role, p.pin, false); return 'Employee added.'; });
}

function api_setPin(p) {
  requireAdmin_();
  checkPinFormat_(p.pin);
  return locked_(() => {
    const e = employeeRows_().find(x => x.username === userKey_(p.email));
    if (!e) throw new Error('Unknown employee.');
    writePin_(e.row, p.pin);
    return 'Password reset for ' + e.name + '.';
  });
}

function api_setActive(p) {
  const me = requireAdmin_();
  return locked_(() => {
    const e = employeeRows_().find(x => x.username === userKey_(p.email));
    if (!e) throw new Error('Unknown employee.');
    if (e.username === me.email && !p.active) throw new Error('You cannot deactivate your own account.');
    sheet_(SH.EMP).getRange(e.row, 5).setValue(p.active ? 'Yes' : 'No');
    return e.name + (p.active ? ' is active.' : ' is deactivated.');
  });
}

function assertOwn_(me, email) {
  if (!me.admin && email !== me.email) throw new Error('You can only change your own records.');
}

function target_(me, email) {
  const e = String(email || me.email).trim().toLowerCase();
  assertOwn_(me, e);
  const t = employees_().find(x => x.email === e);
  if (!t) throw new Error('Unknown employee.');
  return t;
}

/* ---------- Data access ---------- */

function events_() {
  const sh = sheet_(SH.LOG);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 10).getValues().map((r, i) => ({
    row: i + 2, id: r[0], name: r[1], date: r[2], time: r[3], type: r[4], source: r[5],
    email: String(r[7]).toLowerCase(), t: new Date(r[8]).getTime(), status: r[9]
  })).filter(e => e.status === 'Active' && !isNaN(e.t));
}

function newEvent_(person, type, d, source) {
  return { id: 'EV-' + Utilities.getUuid().slice(0, 8).toUpperCase(), name: person.name, email: person.email, type: type, d: d, source: source };
}

function writeEvents_(list) {
  const created = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  appendRows_(SH.LOG, list.map(e => [e.id, e.name, fmtDate_(e.d), fmtTime_(e.d), e.type, e.source, created, e.email, e.d.toISOString(), 'Active']));
}

function audit_(recordId, name, action, prev, upd, by, reason) {
  appendRows_(SH.AUD, [[
    'AU-' + Utilities.getUuid().slice(0, 8).toUpperCase(), recordId, name, action, prev, upd, by,
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), reason
  ]]);
}

/** After adding/moving events, every changed event must sit next to events of the opposite type. */
function checkSequence_(all, email, news, ignoreIds) {
  const list = all.filter(e => e.email === email && ignoreIds.indexOf(e.id) < 0)
    .map(e => ({ id: e.id, type: e.type, t: e.t, isNew: false }));
  news.forEach(n => list.push({ id: n.id, type: n.type, t: n.t, isNew: true }));
  list.sort(cmp_);
  list.forEach((cur, i) => {
    if (!cur.isNew) return;
    [list[i - 1], list[i + 1]].forEach(nb => {
      if (nb && nb.type === cur.type) {
        throw new Error('There is already a ' + cur.type + ' next to ' + Utilities.formatDate(new Date(nb.t), TZ, 'dd/MM HH:mm') +
          '. Entries must alternate between check-in and check-out.');
      }
    });
  });
}

function status_(email) {
  const mine = events_().filter(e => e.email === email).sort(cmp_);
  const last = mine[mine.length - 1];
  if (!last) return { state: 'none', todayMs: 0 };
  const today = ymd_(Date.now());
  let todayMs = 0, open = null;
  mine.forEach(e => {
    if (e.type === 'Check In') open = e;
    else if (open) { if (ymd_(open.t) === today) todayMs += e.t - open.t; open = null; }
  });
  return {
    state: last.type === 'Check In' ? 'in' : 'out', lastType: last.type,
    lastDate: Utilities.formatDate(new Date(last.t), TZ, 'dd/MM/yyyy'), lastTime: fmtTime_(new Date(last.t)),
    sinceMs: last.t, todayMs: todayMs
  };
}

/* ---------- API (called from the page) ---------- */

function api_bootstrap() {
  const me = me_();
  return {
    user: me, nowMs: Date.now(), status: status_(me.email),
    employees: me.admin ? employees_() : []
  };
}

function api_status() {
  const me = me_();
  return { nowMs: Date.now(), status: status_(me.email) };
}

function api_punch(type) {
  if (type !== 'Check In' && type !== 'Check Out') throw new Error('Unknown action.');
  const me = me_();
  return locked_(() => {
    const now = new Date(Math.floor(Date.now() / 60000) * 60000);
    const mine = events_().filter(e => e.email === me.email).sort(cmp_);
    const last = mine[mine.length - 1];
    if (type === 'Check In' && last && last.type === 'Check In') {
      throw new Error('You have been checked in since ' + when_(last) + '. Check out first, or add the missing check-out under Manual entry.');
    }
    if (type === 'Check Out' && (!last || last.type === 'Check Out')) {
      throw new Error('You are not checked in. Check in first, or add the missing check-in under Manual entry.');
    }
    if (last && now.getTime() < last.t) throw new Error('A later registration already exists. Correct it under Manual entry.');
    writeEvents_([newEvent_(me, type, now, 'Automatic')]);
    SpreadsheetApp.flush();
    return {
      message: 'You checked ' + (type === 'Check In' ? 'in' : 'out') + ' successfully at ' + fmtTime_(now) + '.',
      status: status_(me.email)
    };
  });
}

function api_dayEvents(email, ymd) {
  const me = me_();
  const t = target_(me, email);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd || '')) throw new Error('Choose a valid date.');
  return events_().filter(e => e.email === t.email && ymd_(e.t) === ymd).sort(cmp_)
    .map(e => ({ id: e.id, type: e.type, ymd: ymd_(e.t), time: fmtTime_(new Date(e.t)), source: e.source }));
}

/** Add a missed check-in and/or check-out (each optional). */
function api_saveManual(p) {
  const me = me_();
  const t = target_(me, p.email);
  const reason = need_(p.reason);
  return locked_(() => {
    const news = [];
    if (p.inTime) news.push(newEvent_(t, 'Check In', parseLocal_(p.inDate, p.inTime), 'Manual'));
    if (p.outTime) news.push(newEvent_(t, 'Check Out', parseLocal_(p.outDate, p.outTime), 'Manual'));
    if (!news.length) throw new Error('Enter a check-in time, a check-out time, or both.');
    if (news.length === 2 && news[1].d.getTime() <= news[0].d.getTime()) throw new Error('Check-out must be after check-in.');
    news.forEach(n => noFuture_(n.d));
    checkSequence_(events_(), t.email, news.map(n => ({ id: n.id, type: n.type, t: n.d.getTime() })), []);
    writeEvents_(news);
    news.forEach(n => audit_(n.id, t.name, 'Add', '', n.type + ' ' + fmtDate_(n.d) + ' ' + fmtTime_(n.d), me.email, reason));
    SpreadsheetApp.flush();
    return 'Saved ' + news.length + (news.length === 1 ? ' entry.' : ' entries.');
  });
}

function api_updateEvent(p) {
  const me = me_();
  const reason = need_(p.reason);
  return locked_(() => {
    const all = events_();
    const ev = all.find(e => e.id === p.id);
    if (!ev) throw new Error('Record not found.');
    assertOwn_(me, ev.email);
    const d = parseLocal_(p.date, p.time);
    noFuture_(d);
    if (d.getTime() === ev.t) throw new Error('Nothing changed.');
    checkSequence_(all, ev.email, [{ id: ev.id, type: ev.type, t: d.getTime() }], [ev.id]);
    const sh = sheet_(SH.LOG);
    sh.getRange(ev.row, 3, 1, 4).setValues([[fmtDate_(d), fmtTime_(d), ev.type, ev.source === 'Automatic' ? 'Edited' : ev.source]]);
    sh.getRange(ev.row, 9).setValue(d.toISOString());
    audit_(ev.id, ev.name, 'Edit', ev.type + ' ' + ev.date + ' ' + ev.time, ev.type + ' ' + fmtDate_(d) + ' ' + fmtTime_(d), me.email, reason);
    SpreadsheetApp.flush();
    return 'Entry updated.';
  });
}

function api_deleteEvent(p) {
  const me = me_();
  const reason = need_(p.reason);
  return locked_(() => {
    const ev = events_().find(e => e.id === p.id);
    if (!ev) throw new Error('Record not found.');
    assertOwn_(me, ev.email);
    sheet_(SH.LOG).getRange(ev.row, 10).setValue('Deleted');
    audit_(ev.id, ev.name, 'Delete', ev.type + ' ' + ev.date + ' ' + ev.time, '', me.email, reason);
    SpreadsheetApp.flush();
    return 'Entry deleted. The original stays in the sheet marked Deleted.';
  });
}

/** Pairs check-ins with check-outs and totals real elapsed time per day. */
function api_report(p) {
  const me = me_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.from || '') || !/^\d{4}-\d{2}-\d{2}$/.test(p.to || '')) throw new Error('Choose a valid From and To date.');
  if (p.from > p.to) throw new Error('The From date must be on or before the To date.');
  const all = events_();
  let emails;
  if (me.admin) {
    emails = p.email && p.email !== 'all' ? [String(p.email).toLowerCase()]
      : Array.from(new Set(all.map(e => e.email)));
  } else {
    emails = [me.email];
  }

  const days = {};
  let incomplete = 0;

  emails.forEach(email => {
    const evs = all.filter(e => e.email === email).sort(cmp_);
    const items = [];
    let open = null;
    evs.forEach(e => {
      if (e.type === 'Check In') {
        if (open) items.push({ k: 'noOut', inE: open });
        open = e;
      } else if (open) {
        items.push({ k: 'ok', inE: open, outE: e });
        open = null;
      } else {
        items.push({ k: 'noIn', outE: e });
      }
    });
    if (open) items.push({ k: Date.now() - open.t < 20 * 3600000 ? 'live' : 'noOut', inE: open });

    items.forEach(it => {
      const base = it.inE || it.outE;
      const ymd = ymd_(base.t);
      if (ymd < p.from || ymd > p.to) return;
      const key = email + '|' + ymd;
      const d = days[key] || (days[key] = { ymd: ymd, email: email, name: base.name, ms: 0, inT: null, outT: null, flags: {}, sessions: [], manual: false });
      if (it.k === 'noOut' || it.k === 'noIn') { d.flags.incomplete = true; incomplete++; }
      if (it.k === 'live') d.flags.live = true;
      const s = { k: it.k };
      if (it.inE) {
        s.inId = it.inE.id; s.inTime = fmtTime_(new Date(it.inE.t));
        if (d.inT === null || it.inE.t < d.inT) d.inT = it.inE.t;
        if (it.inE.source !== 'Automatic') d.manual = true;
      }
      if (it.outE) {
        s.outId = it.outE.id; s.outTime = fmtTime_(new Date(it.outE.t));
        if (ymd_(it.outE.t) !== ymd) s.outTime += ' (+1 day)';
        if (d.outT === null || it.outE.t > d.outT) d.outT = it.outE.t;
        if (it.outE.source !== 'Automatic') d.manual = true;
      }
      if (it.k === 'ok') { const ms = it.outE.t - it.inE.t; d.ms += ms; s.hours = fmtDur_(ms); }
      d.sessions.push(s);
    });
  });

  let totalMs = 0, worked = 0;
  const rows = Object.keys(days).map(k => days[k]).sort((a, b) => a.ymd < b.ymd ? -1 : a.ymd > b.ymd ? 1 : a.name.localeCompare(b.name))
    .map(d => {
      totalMs += d.ms;
      if (d.ms > 0) worked++;
      let out = '—';
      if (d.outT !== null) out = fmtTime_(new Date(d.outT)) + (ymd_(d.outT) !== d.ymd ? ' (+1 day)' : '');
      return {
        ymd: d.ymd, date: dmy_(d.ymd), email: d.email, name: d.name,
        checkIn: d.inT !== null ? fmtTime_(new Date(d.inT)) : '—', checkOut: out,
        hours: fmtDur_(d.ms), manual: d.manual, sessions: d.sessions,
        status: d.flags.incomplete ? 'Incomplete' : d.flags.live ? 'In progress' : 'Complete'
      };
    });

  return {
    days: rows,
    summary: {
      total: fmtDur_(totalMs), daysWorked: worked,
      average: worked ? fmtDur_(totalMs / worked) : '0h 00m', incomplete: incomplete
    }
  };
}
