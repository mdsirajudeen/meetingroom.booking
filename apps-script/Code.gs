/*********************************************************************
 * Kuwait Hospital – Meeting Room Booking System (Google Apps Script)
 * Staff book a room with the form -> room coordinator approves/rejects
 * (or the room is auto-approved) -> requester gets an email ->
 * the room tablet page shows Occupied / Available live.
 *
 * GITHUB PAGES VERSION: the web pages are hosted on GitHub Pages and talk to
 * this script through doPost (a small JSON API). This project needs only
 * Code.gs and Logo.gs – no HTML files.
 *********************************************************************/

/* ---------- SETTINGS (edit these) ---------- */
const TZ               = 'Asia/Kuwait';
const TZ_OFFSET        = '+03:00';          // Kuwait has no daylight saving
const SLOT_MINUTES     = 30;                // bookings start/end on these steps
const DAY_START        = '07:00';
const DAY_END          = '17:00';
const WORK_DAYS        = [0, 1, 2, 3, 4];   // 0 = Sunday ... 4 = Thursday
const MAX_DAYS_AHEAD   = 60;
const MAX_HOURS        = 8;                 // longest single booking
const ALLOWED_DOMAIN   = 'kuwaithospital.com.kw';   // only these requester emails can book ('' = anyone)
const HOSPITAL_NAME    = 'Kuwait Hospital';
const SENDER_NAME      = HOSPITAL_NAME + ' Meeting Rooms';
const MAIN_COORDINATOR_EMAIL = 'msirajudeen@kuwaithospital.com.kw';   // gets requests for rooms without a coordinator
const IT_SUPPORT_EMAIL  = 'itsupport@kuwaithospital.com.kw';    // gets IT needs (laptop, projector...) when a booking is approved / cancelled
const HOSPITALITY_EMAIL = 'hospitality@kuwaithospital.com.kw';  // gets refreshment orders when a booking is approved / cancelled
const SESSION_HOURS    = 6;
const WEB_APP_URL      = '';                // this project's .../exec link (the API address)
const SITE_URL         = '';                // your GitHub Pages address, e.g. 'https://yourname.github.io/kh-meeting-rooms/'
const LOGO_FILE_ID     = '';                // leave '' to use Logo.gs

/* ---------- Sheet layout ---------- */
const BOOK = 'Bookings', ROOMS = 'Rooms', COORDS_TAB = 'Coordinators';
const DEFAULT_ROOMS = [                     // first-time list; after setup edit the Rooms tab
  ['Board Room', 20, 'Admin floor', 'Yes'],
  ['Conference Room', 12, '', 'Yes'],
  ['Training Room', 30, '', 'Yes']
];
const HEADERS = ['Booking ID', 'Submitted', 'Status', 'Room', 'Date', 'Start', 'End', 'Meeting Title', 'Requester',
                 'Email', 'Department', 'Mobile', 'Attendees', 'IT Requirements', 'Private', 'Notes',
                 'Decided', 'Decided By', 'Coordinator Note', 'Refreshments'];
const COL = { id: 0, submitted: 1, status: 2, room: 3, date: 4, start: 5, end: 6, title: 7, name: 8, email: 9,
              dept: 10, mobile: 11, attendees: 12, needs: 13, priv: 14, notes: 15, decided: 16, decidedBy: 17, cnote: 18, refresh: 19 };
const BLOCKING = ['Pending', 'Approved'];   // these statuses hold the room

function setup() {
  const ss = SpreadsheetApp.getActive();
  const bk = ss.getSheetByName(BOOK) || ss.insertSheet(BOOK);
  const isNew = bk.getLastRow() === 0;
  bk.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  if (isNew) bk.setFrozenRows(1);

  let rm = ss.getSheetByName(ROOMS);
  if (!rm) {
    rm = ss.insertSheet(ROOMS);
    rm.getRange(1, 1, 1, 5).setValues([['Room', 'Capacity', 'Location', 'Needs approval (Yes/No)', 'Active (Yes/No)']]).setFontWeight('bold');
    rm.getRange(2, 1, DEFAULT_ROOMS.length, 5).setValues(DEFAULT_ROOMS.map(r => r.concat(['Yes'])));
    rm.setFrozenRows(1); rm.setColumnWidths(1, 5, 180);
  }
  let co = ss.getSheetByName(COORDS_TAB);
  if (!co) {
    co = ss.insertSheet(COORDS_TAB);
    co.getRange(1, 1, 1, 4).setValues([['Username', 'Rooms (comma separated, or ALL)', 'Notification Email', 'Active (Yes/No)']])
      .setFontWeight('bold');
    co.setFrozenRows(1); co.setColumnWidths(1, 4, 220);
  }
  Logger.log('API (paste into js/config.js): ' + baseUrl_());
  Logger.log('Booking form: ' + formUrl_());
  Logger.log('Coordinator portal: ' + portalUrl_());
  Logger.log('Next: reload the Sheet and use the "Meeting Rooms" menu to add coordinators.');
}

/* ---------- Rooms ---------- */
function rooms_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ROOMS);
  const rows = sh ? sh.getDataRange().getValues().slice(1) : DEFAULT_ROOMS.map(r => r.concat(['Yes']));
  return rows.filter(r => String(r[0]).trim() && String(r[4]).trim().toLowerCase() !== 'no')
    .map(r => ({ name: String(r[0]).trim(), capacity: Number(r[1]) || 0, location: String(r[2] || '').trim(),
                 approval: String(r[3]).trim().toLowerCase() !== 'no' }))
    .sort((a, b) => a.name.localeCompare(b.name, 'en', { sensitivity: 'base' }));
}
function getRooms() { return rooms_(); }
function room_(name) { return rooms_().find(r => norm_(r.name) === norm_(name)) || null; }

/* ---------- Web entry points ---------- */
function doGet() {
  const site = siteUrl_();
  return HtmlService.createHtmlOutput('<p style="font-family:Arial;margin:40px">Kuwait Hospital meeting room booking: ' +
    (site ? '<a href="' + esc_(site) + '" target="_top">open the booking site</a>' : 'API is running.') + '</p>').setTitle('Meeting Rooms');
}
/* JSON API used by the GitHub Pages site (js/api.js). Only these functions can be called. */
const API_FUNCTIONS = ['getRooms', 'getHours', 'getBusy', 'submitBooking', 'getRoomStatus', 'coordLogin', 'coordLogout',
  'coordWhoAmI', 'coordChangePassword', 'coordList', 'coordDecide', 'coordUsers', 'coordResetPassword'];
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (API_FUNCTIONS.indexOf(req.fn) < 0) throw new Error('Unknown request.');
    out = { ok: true, data: globalThis[req.fn].apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Booking form ---------- */
/* Busy intervals for a room and date: [{start:'09:00', end:'10:30'}] (Pending + Approved) */
function getBusy(room, date) {
  checkDate_(date);
  const r = room_(room);
  if (!r) throw new Error('Please choose a room from the list.');
  return readBookings_().filter(x => norm_(x.r[COL.room]) === norm_(r.name) && isoOf_(x.r[COL.date]) === date &&
                                     BLOCKING.includes(String(x.r[COL.status])))
    .map(x => ({ start: hhmm_(x.r[COL.start]), end: hhmm_(x.r[COL.end]) }));
}
function getHours() { return { start: DAY_START, end: DAY_END, step: SLOT_MINUTES, maxHours: MAX_HOURS, days: WORK_DAYS,
                               ahead: MAX_DAYS_AHEAD, domain: ALLOWED_DOMAIN }; }

function submitBooking(f) {
  f = f || {};
  const v = (k, max) => String(f[k] || '').trim().slice(0, max || 200);
  const d = { name: v('name'), email: v('email').toLowerCase(), dept: v('dept'), mobile: v('mobile', 20), room: v('room'),
              title: v('title'), date: v('date', 10), start: v('start', 5), end: v('end', 5),
              attendees: Math.max(0, parseInt(f.attendees, 10) || 0), needs: v('needs', 300), refresh: v('refresh', 400),
              priv: f.priv === true || f.priv === 'on' || f.priv === 'yes', notes: v('notes', 1000) };
  if (!d.name || !d.email || !d.dept || !d.room || !d.title || !d.date || !d.start || !d.end || !d.attendees)
    throw new Error('Please fill in all required fields.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) throw new Error('Please enter a valid email address.');
  if (ALLOWED_DOMAIN && !d.email.endsWith('@' + ALLOWED_DOMAIN))
    throw new Error('Please use your hospital email address (@' + ALLOWED_DOMAIN + ').');
  if (d.mobile && !/^[0-9+\-\s]{7,15}$/.test(d.mobile)) throw new Error('Please enter a valid mobile number, or leave it empty.');
  const room = room_(d.room);
  if (!room) throw new Error('Please choose a room from the list.');
  checkDate_(d.date);
  const s = toMin_(d.start), en = toMin_(d.end);
  if (!(s >= toMin_(DAY_START) && en <= toMin_(DAY_END) && en > s && (s - toMin_(DAY_START)) % SLOT_MINUTES === 0 &&
        (en - toMin_(DAY_START)) % SLOT_MINUTES === 0)) throw new Error('Please choose a valid start and end time.');
  if (en - s > MAX_HOURS * 60) throw new Error('A booking can be at most ' + MAX_HOURS + ' hours.');
  if (room.capacity && d.attendees > room.capacity)
    throw new Error(room.name + ' holds up to ' + room.capacity + ' people. Please choose a larger room.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const id = 'RB-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
  const status = room.approval ? 'Pending' : 'Approved';
  try {
    if (d.date === today_() && s <= toMin_(Utilities.formatDate(new Date(), TZ, 'HH:mm')))
      throw new Error('That start time has already passed today.');
    const clash = getBusy(room.name, d.date).find(b => s < toMin_(b.end) && en > toMin_(b.start));
    if (clash) throw new Error(room.name + ' is already booked ' + clash.start + ' – ' + clash.end + '. Please choose another time.');
    const row = [id, Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), status, room.name, dmy_(d.date), d.start, d.end,
                 d.title, d.name, d.email, d.dept, d.mobile, String(d.attendees), d.needs, d.priv ? 'Yes' : 'No', d.notes,
                 room.approval ? '' : Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), room.approval ? '' : 'Auto-approved', '', d.refresh];
    const sh = sheet_(BOOK);
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  const o = { id, room: room.name, when: prettyDate_(d.date) + ', ' + d.start + ' – ' + d.end, title: d.title, name: d.name,
              email: d.email, dept: d.dept, mobile: d.mobile, attendees: d.attendees, needs: d.needs, refresh: d.refresh,
              priv: d.priv, notes: d.notes, location: room.location };
  if (room.approval) {
    send_({ to: roomEmails_(room.name), replyTo: d.email, name: SENDER_NAME,
      subject: 'Room request: ' + room.name + ', ' + o.when + ' – ' + d.name,
      htmlBody: emailHtml_('New room booking request', 'Please review this booking and approve or reject it. The room is held until you respond.',
        details_(o), btn_(portalUrl_() + '?id=' + encodeURIComponent(id), 'Review and respond', BRAND_BLUE) +
        '<p style="margin:14px 0 0;font-size:13px;color:#5b6780">You will be asked to sign in.</p>') });
    send_({ to: d.email, name: SENDER_NAME, subject: 'Booking request received: ' + room.name + ', ' + o.when,
      htmlBody: emailHtml_('Booking request received', 'Your request has been sent for approval. You will receive another email once it is approved or rejected.', details_(o)) });
  } else {
    send_({ to: d.email, name: SENDER_NAME, subject: 'Room booked: ' + room.name + ', ' + o.when,
      htmlBody: emailHtml_('Room booked', 'Your booking is confirmed.', details_(o)) });
    notifyServices_(o, 'Approved');
  }
  return { id, status, when: o.when, room: room.name };
}

/* ---------- Room tablet display ---------- */
function getRoomStatus(room) {
  const r = room_(room);
  if (!r) throw new Error('Unknown room: ' + room);
  const today = today_();
  const list = readBookings_().filter(x => norm_(x.r[COL.room]) === norm_(r.name) && isoOf_(x.r[COL.date]) === today &&
                                           String(x.r[COL.status]) === 'Approved')
    .map(x => { const priv = String(x.r[COL.priv]).toLowerCase() === 'yes';
      return { start: hhmm_(x.r[COL.start]), end: hhmm_(x.r[COL.end]),
               startMs: Date.parse(today + 'T' + hhmm_(x.r[COL.start]) + ':00' + TZ_OFFSET),
               endMs: Date.parse(today + 'T' + hhmm_(x.r[COL.end]) + ':00' + TZ_OFFSET),
               title: priv ? 'Private meeting' : String(x.r[COL.title]), host: priv ? '' : String(x.r[COL.name]),
               dept: priv ? '' : String(x.r[COL.dept]) }; })
    .sort((a, b) => a.startMs - b.startMs);
  return { room: r.name, capacity: r.capacity, location: r.location, now: Date.now(), today: prettyDate_(today), meetings: list };
}

/* ---------- Coordinator roles (Coordinators tab) ---------- */
function roles_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(COORDS_TAB), out = {};
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const u = String(r[0]).trim().toLowerCase();
    if (!u) return;
    const raw = String(r[1]).split(',').map(x => x.trim()).filter(Boolean);
    out[u] = { all: raw.some(x => x.toUpperCase() === 'ALL'), rooms: raw.filter(x => x.toUpperCase() !== 'ALL'),
               email: String(r[2]).trim(), active: String(r[3]).trim().toLowerCase() !== 'no' };
  });
  return out;
}
function handlers_(room, roles) {
  const r = roles || roles_();
  return Object.keys(r).filter(u => r[u].active && !r[u].all && r[u].rooms.some(x => norm_(x) === norm_(room)));
}
function roomEmails_(room) {
  const r = roles_(), list = handlers_(room, r).map(u => r[u].email).filter(Boolean);
  return list.length ? Array.from(new Set(list)).join(',') : (MAIN_COORDINATOR_EMAIL || Session.getEffectiveUser().getEmail());
}
function access_(token, allowTemp) {
  const user = session_(token), roles = roles_(), role = roles[user];
  if (!role || !role.active || (!role.all && !role.rooms.length)) throw new Error('NO_ACCESS');
  if (!allowTemp && mustChange_(user)) throw new Error('MUST_CHANGE');
  return { user, all: role.all, clinics: role.rooms, roles, label: role.all ? 'Supervisor (all rooms)' : role.rooms.join(', ') };
}
function canSee_(acc, room) { return acc.all || handlers_(room, acc.roles).indexOf(acc.user) >= 0; }
function baseUrl_() { return WEB_APP_URL || ScriptApp.getService().getUrl(); }
function siteUrl_() { return SITE_URL ? SITE_URL.replace(/\/?$/, '/') : ''; }
function formUrl_() { return siteUrl_() || baseUrl_(); }
function portalUrl_() { return siteUrl_() + 'coordinator.html'; }
function roomUrl_(name) { return siteUrl_() + 'room.html?room=' + encodeURIComponent(name); }

/* ---------- Coordinator portal ---------- */
function coordList(token) {
  const acc = access_(token);
  return readBookings_().map(x => {
    const r = x.r, date = isoOf_(r[COL.date]);
    return { id: String(r[COL.id]), status: String(r[COL.status]), clinic: String(r[COL.room]),
             when: prettyDate_(date) + ', ' + hhmm_(r[COL.start]) + ' – ' + hhmm_(r[COL.end]), date, time: hhmm_(r[COL.start]),
             title: String(r[COL.title]), name: String(r[COL.name]), email: String(r[COL.email]), dept: String(r[COL.dept]),
             mobile: String(r[COL.mobile]), attendees: String(r[COL.attendees]), needs: String(r[COL.needs]),
             refresh: String(r[COL.refresh] || ''),
             priv: String(r[COL.priv]).toLowerCase() === 'yes', notes: String(r[COL.notes]), submitted: String(r[COL.submitted]),
             note: String(r[COL.cnote] || ''), decidedBy: String(r[COL.decidedBy] || ''), future: date >= today_() };
  }).filter(o => canSee_(acc, o.clinic))
    .sort((a, b) => (a.status === 'Pending' ? 0 : 1) - (b.status === 'Pending' ? 0 : 1) ||
                    (a.status === 'Pending' ? 1 : -1) * (a.date + a.time).localeCompare(b.date + b.time));
}
/* action: approve | reject | cancel (cancel = an approved booking is called off, the room becomes free) */
function coordDecide(token, id, action, note) {
  const acc = access_(token);
  note = String(note || '').trim().slice(0, 500);
  if (['approve', 'reject', 'cancel'].indexOf(action) < 0) throw new Error('Unknown action.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let rec;
  try {
    rec = readBookings_().find(x => String(x.r[COL.id]) === String(id));
    if (!rec) throw new Error('Booking not found.');
    if (!canSee_(acc, rec.r[COL.room])) throw new Error('This booking belongs to another room.');
    const cur = String(rec.r[COL.status]);
    if (action === 'cancel' ? cur !== 'Approved' : cur !== 'Pending')
      throw new Error('This booking is already ' + cur.toLowerCase() + '.');
    const status = { approve: 'Approved', reject: 'Rejected', cancel: 'Cancelled' }[action];
    const sh = sheet_(BOOK);
    sh.getRange(rec.row, COL.status + 1).setValue(status);
    sh.getRange(rec.row, COL.decided + 1, 1, 3).setNumberFormat('@')
      .setValues([[Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), acc.user, note]]);
    rec.r[COL.status] = status;
  } finally { lock.releaseLock(); }
  const r = rec.r, date = isoOf_(r[COL.date]);
  const o = { id: String(r[COL.id]), room: String(r[COL.room]), when: prettyDate_(date) + ', ' + hhmm_(r[COL.start]) + ' – ' + hhmm_(r[COL.end]),
              title: String(r[COL.title]), name: String(r[COL.name]), email: String(r[COL.email]), dept: String(r[COL.dept]),
              mobile: String(r[COL.mobile]), attendees: String(r[COL.attendees]), needs: String(r[COL.needs]),
              refresh: String(r[COL.refresh] || ''), priv: String(r[COL.priv]).toLowerCase() === 'yes', notes: String(r[COL.notes]),
              location: (room_(r[COL.room]) || {}).location || '' };
  const st = String(r[COL.status]);
  const txt = { Approved: ['Room booking approved', 'Your booking is confirmed.'],
                Rejected: ['Room booking not approved', 'Unfortunately this booking could not be approved. Please choose another time or room.'],
                Cancelled: ['Room booking cancelled', 'Your approved booking has been cancelled by the coordinator.'] }[st];
  send_({ to: o.email, replyTo: roomEmails_(o.room).split(',')[0], name: SENDER_NAME,
          subject: st + ': ' + o.room + ', ' + o.when,
          htmlBody: emailHtml_(txt[0], txt[1], details_(o).concat(note ? [['Message from coordinator', note]] : [])) });
  notifyServices_(o, st);                              // IT support / hospitality: set up on approval, stand down on cancellation
  return { status: st };
}

/* ---------- Sign-in, sessions and passwords ---------- */
function setPassword_(user, pw, mustChange) {
  const props = PropertiesService.getScriptProperties(), salt = Utilities.getUuid();
  props.setProperty('USER_' + user, salt + ':' + hash_(salt, pw));
  if (mustChange) props.setProperty('MUSTCHG_' + user, '1'); else props.deleteProperty('MUSTCHG_' + user);
  props.setProperty('UEP_' + user, String(Date.now()));
}
function tempPassword_() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789'; let pw = '';
  for (let i = 0; i < 10; i++) pw += c.charAt(Math.floor(Math.random() * c.length));
  return pw;
}
function checkNewPassword_(pw) {
  if (String(pw || '').length < 8) throw new Error('The new password must be at least 8 characters.');
  if (!/[A-Za-z]/.test(pw) || !/[0-9]/.test(pw)) throw new Error('Use at least one letter and one number.');
}
function hash_(salt, pw) {
  let h = salt + '|' + pw;
  for (let i = 0; i < 300; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  return h;
}
function signOutAll_() { PropertiesService.getScriptProperties().setProperty('SESSION_EPOCH', String(Date.now())); }
function sessKey_(t) { return 's' + (PropertiesService.getScriptProperties().getProperty('SESSION_EPOCH') || '0') + '_' + t; }
function session_(token) {
  const cache = CacheService.getScriptCache(), key = token ? sessKey_(String(token)) : '';
  const val = key ? cache.get(key) : null, props = PropertiesService.getScriptProperties();
  if (!val) throw new Error('SESSION_EXPIRED');
  const [user, ep] = val.split('|');
  if (!props.getProperty('USER_' + user) || (ep || '0') !== (props.getProperty('UEP_' + user) || '0')) throw new Error('SESSION_EXPIRED');
  cache.put(key, val, Math.min(SESSION_HOURS, 6) * 3600);
  return user;
}
function newSession_(user) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const ep = PropertiesService.getScriptProperties().getProperty('UEP_' + user) || '0';
  CacheService.getScriptCache().put(sessKey_(token), user + '|' + ep, Math.min(SESSION_HOURS, 6) * 3600);
  return token;
}
function mustChange_(user) { return PropertiesService.getScriptProperties().getProperty('MUSTCHG_' + user) === '1'; }
function info_(acc) { return { user: acc.user, label: acc.label, all: acc.all, clinics: acc.clinics, mustChange: mustChange_(acc.user) }; }
function coordLogin(user, password) {
  user = String(user || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), failKey = 'fail_' + user, fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('Too many wrong attempts. Please wait 15 minutes and try again.');
  const stored = PropertiesService.getScriptProperties().getProperty('USER_' + user), parts = stored ? stored.split(':') : [];
  if (!stored || hash_(parts[0], String(password || '')) !== parts[1]) {
    cache.put(failKey, String(fails + 1), 900); Utilities.sleep(700); throw new Error('Wrong username or password.');
  }
  cache.remove(failKey);
  const role = roles_()[user];
  if (!role || !role.active || (!role.all && !role.rooms.length))
    throw new Error('This account is not assigned to a room, or is inactive. Please contact the administrator.');
  const token = newSession_(user);
  return Object.assign({ token }, info_(access_(token, true)));
}
function coordLogout(token) { if (token) CacheService.getScriptCache().remove(sessKey_(String(token))); }
function coordWhoAmI(token) { return info_(access_(token, true)); }
function coordChangePassword(token, current, next) {
  const acc = access_(token, true), parts = (PropertiesService.getScriptProperties().getProperty('USER_' + acc.user) || '').split(':');
  if (hash_(parts[0], String(current || '')) !== parts[1]) { Utilities.sleep(700); throw new Error('Your current password is not correct.'); }
  next = String(next || ''); checkNewPassword_(next);
  if (next === String(current)) throw new Error('The new password must be different from the current one.');
  setPassword_(acc.user, next, false);
  const t = newSession_(acc.user);
  return Object.assign({ token: t }, info_(access_(t)));
}
function supervisor_(token) { const acc = access_(token); if (!acc.all) throw new Error('Only a supervisor can do this.'); return acc; }
function coordUsers(token) {
  const acc = supervisor_(token), r = acc.roles;
  return Object.keys(r).sort().map(u => ({ user: u, label: r[u].all ? 'Supervisor (all rooms)' : r[u].rooms.join(', '),
    email: r[u].email, active: r[u].active, mustChange: mustChange_(u), me: u === acc.user }));
}
function coordResetPassword(token, user) {
  const acc = supervisor_(token); user = String(user || '').trim().toLowerCase();
  if (user === acc.user) throw new Error('Use "Change my password" for your own account.');
  if (!acc.roles[user]) throw new Error('No coordinator with that username.');
  const pw = tempPassword_(); setPassword_(user, pw, true);
  return { user, password: pw };
}

/* ---------- Google Sheet menu ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Meeting Rooms')
    .addItem('Add / update coordinator', 'menuAddCoordinator')
    .addItem('Reset coordinator password', 'menuSetPassword')
    .addItem('Remove a coordinator', 'menuRemoveCoordinator')
    .addItem('Sign out everyone', 'menuSignOutAll')
    .addSeparator()
    .addItem('Show links (form, portal, room tablets)', 'menuShowLinks')
    .addToUi();
}
function askText_(title, text, validate) {
  const ui = SpreadsheetApp.getUi();
  for (;;) {
    const r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return null;
    const v = r.getResponseText().trim(), err = validate ? validate(v) : '';
    if (!err) return v;
    ui.alert(err);
  }
}
function menuAddCoordinator() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  setup();
  const user = askText_('Coordinator (1/4)', 'Username (e.g. rooms1). Letters, numbers, dot, dash, underscore or @.',
    v => /^[a-z0-9._@-]{3,60}$/i.test(v) ? '' : 'Username must be 3-60 characters with no spaces.');
  if (user === null) return;
  const u = user.toLowerCase(), names = rooms_().map(r => r.name), existing = roles_()[u];
  const rm = askText_('Coordinator (2/4)', 'Room(s) ' + u + ' will handle, separated by commas, or ALL for a supervisor.\n\nRooms: ' +
    names.join(', ') + (existing ? '\n\nCurrently: ' + (existing.all ? 'ALL' : existing.rooms.join(', ')) : ''),
    v => { const p = v.split(',').map(x => x.trim()).filter(Boolean);
           if (!p.length) return 'Please enter at least one room, or ALL.';
           const bad = p.filter(x => x.toUpperCase() !== 'ALL' && !room_(x));
           return bad.length ? 'Not in the Rooms tab: ' + bad.join(', ') : ''; });
  if (rm === null) return;
  const parts = rm.split(',').map(x => x.trim()).filter(Boolean);
  const roomText = parts.some(x => x.toUpperCase() === 'ALL') ? 'ALL' : parts.map(x => room_(x).name).join(', ');
  const email = askText_('Coordinator (3/4)', 'Email that receives booking requests for these rooms (optional):' +
    (existing && existing.email ? '\n\nCurrently: ' + existing.email : ''),
    v => !v || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) ? '' : 'Please enter a valid email address, or leave it blank.');
  if (email === null) return;
  const hasPw = !!props.getProperty('USER_' + u);
  const pw = askText_('Coordinator (4/4)', 'Temporary password for ' + u + ' (at least 8 characters). They change it at first sign-in' +
    (hasPw ? '.\nLeave blank to keep the current password.' : '.'), v => (!v && hasPw) || v.length >= 8 ? '' : 'At least 8 characters.');
  if (pw === null) return;
  if (pw) setPassword_(u, pw, true);
  const sh = SpreadsheetApp.getActive().getSheetByName(COORDS_TAB), vals = sh.getDataRange().getValues();
  let row = vals.findIndex((r, i) => i > 0 && String(r[0]).trim().toLowerCase() === u) + 1;
  if (!row) row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 4).setNumberFormat('@').setValues([[u, roomText, email || (existing ? existing.email : ''), 'Yes']]);
  ui.alert('Saved.\n\nUsername: ' + u + '\nRooms: ' + roomText + '\n\nPortal:\n' + portalUrl_());
}
function menuSetPassword() {
  const ui = SpreadsheetApp.getUi();
  const u = askText_('Reset password', 'Username:', v => roles_()[v.toLowerCase()] ? '' : 'No coordinator with that username.');
  if (u === null) return;
  const pw = askText_('Reset password', 'Temporary password for ' + u.toLowerCase() + ' (at least 8 characters):',
    v => v.length >= 8 ? '' : 'At least 8 characters.');
  if (pw === null) return;
  setPassword_(u.toLowerCase(), pw, true);
  ui.alert('Temporary password set. ' + u.toLowerCase() + ' must choose a new password at next sign-in.');
}
function menuRemoveCoordinator() {
  const ui = SpreadsheetApp.getUi(), u = askText_('Remove coordinator', 'Username to remove:', null);
  if (u === null) return;
  const key = u.toLowerCase(), sh = SpreadsheetApp.getActive().getSheetByName(COORDS_TAB);
  PropertiesService.getScriptProperties().deleteProperty('USER_' + key);
  if (sh) { const v = sh.getDataRange().getValues();
            for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]).trim().toLowerCase() === key) sh.deleteRow(i + 1); }
  signOutAll_();
  ui.alert(key + ' removed; everyone has been signed out.');
}
function menuSignOutAll() { signOutAll_(); SpreadsheetApp.getUi().alert('Everyone has been signed out.'); }
function menuShowLinks() {
  const lines = rooms_().map(r => r.name + ':\n' + roomUrl_(r.name)).join('\n\n');
  SpreadsheetApp.getUi().alert('Links', 'Booking form:\n' + formUrl_() + '\n\nCoordinator portal:\n' + portalUrl_() +
    '\n\nRoom tablet screens:\n' + lines, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ---------- Helpers ---------- */
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found. Run setup() first.');
  return sh;
}
function readBookings_() { return sheet_(BOOK).getDataRange().getValues().slice(1).map((r, i) => ({ row: i + 2, r })); }
function norm_(v) { return String(v || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function hhmm_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v).trim(); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function toMin_(s) { const [h, m] = String(s).split(':').map(Number); return h * 60 + m; }
function isoOf_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const t = String(v).trim(), m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : t;
}
function dmy_(iso) { const [y, m, d] = iso.split('-'); return d + '/' + m + '/' + y; }
function prettyDate_(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][new Date(Date.UTC(y, m - 1, d)).getUTCDay()] + ' ' + dmy_(iso);
}
function checkDate_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date))) throw new Error('Please choose a date.');
  const [y, m, d] = date.split('-').map(Number);
  if (!WORK_DAYS.includes(new Date(Date.UTC(y, m - 1, d)).getUTCDay()))
    throw new Error('Rooms can be booked on ' + WORK_DAYS.map(n => ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][n]).join(', ') + ' only.');
  if (date < today_()) throw new Error('Please choose today or a later date.');
  const max = Utilities.formatDate(new Date(Date.now() + MAX_DAYS_AHEAD * 864e5), TZ, 'yyyy-MM-dd');
  if (date > max) throw new Error('Please choose a date within the next ' + MAX_DAYS_AHEAD + ' days.');
}
function details_(o) {
  return [['Reference', o.id], ['Room', o.room], ['Date and time', o.when], ['Meeting', o.title + (o.priv ? ' (private)' : '')],
          ['Requested by', o.name], ['Department', o.dept], ['Email', o.email]]
    .concat(o.mobile ? [['Mobile', o.mobile]] : [])
    .concat([['Attendees', String(o.attendees)], ['IT requirements', o.needs || '-'], ['Refreshments', o.refresh || '-']])
    .concat(o.notes ? [['Notes', o.notes]] : []);
}
/* IT support and hospitality get their own email when a booking is approved (set up) or cancelled (stand down).
   Nothing is sent while a booking is pending or if it is rejected. */
function notifyServices_(o, status) {
  if (status !== 'Approved' && status !== 'Cancelled') return;
  const cancelled = status === 'Cancelled';
  const base = [['Reference', o.id], ['Room', o.room + (o.location ? ' – ' + o.location : '')], ['Date and time', o.when],
                ['Meeting', o.priv ? 'Private meeting' : o.title], ['Requested by', o.name + ' (' + o.dept + ')'],
                ['Contact', [o.mobile, o.email].filter(Boolean).join(' · ')], ['Attendees', String(o.attendees)]];
  const jobs = [[IT_SUPPORT_EMAIL, o.needs, 'IT setup', 'IT requirements'],
                [HOSPITALITY_EMAIL, o.refresh, 'Refreshments', 'Refreshments']];
  jobs.forEach(([to, items, what, label]) => {
    if (!to || !items) return;
    send_({ to, replyTo: o.email, name: SENDER_NAME,
      subject: (cancelled ? 'CANCELLED – ' : '') + what + ': ' + o.room + ', ' + o.when,
      htmlBody: emailHtml_(cancelled ? what + ' no longer needed' : what + ' request',
        cancelled ? 'This meeting has been cancelled. Please do not prepare the items below.'
                  : 'A meeting room booking has been approved. Please prepare the following before the meeting starts.',
        base.concat([[label, items]]).concat(o.notes && !cancelled ? [['Notes', o.notes]] : [])) });
  });
}

/* ---------- Email ---------- */
const BRAND_BLUE = '#0b4a8f';
function logoBlob_() {
  if (LOGO_FILE_ID) { try { return DriveApp.getFileById(LOGO_FILE_ID).getBlob(); } catch (e) {} }
  if (typeof LOGO_PNG_BASE64 !== 'undefined' && LOGO_PNG_BASE64)
    return Utilities.newBlob(Utilities.base64Decode(LOGO_PNG_BASE64), 'image/png', 'logo.png');
  return null;
}
function logoDataUri_() {
  const cache = CacheService.getScriptCache(), key = 'logo_' + LOGO_FILE_ID, hit = cache.get(key);
  if (hit !== null) return hit;
  const b = logoBlob_(), uri = b ? 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) : '';
  try { cache.put(key, uri, 21600); } catch (e) {}
  return uri;
}
function send_(opts) {
  const b = logoBlob_();
  if (b) opts.inlineImages = { logo: b };
  if (!opts.body) opts.body = textOf_(opts.htmlBody || '');
  MailApp.sendEmail(opts);
}
function textOf_(html) {
  return String(html)
    .replace(/<a [^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, (m, h, t) => t.replace(/<[^>]+>/g, '') + ': ' + h.replace(/&amp;/g, '&') + '\n')
    .replace(/<\/(tr|p|h2|div)>/gi, '\n').replace(/<\/td><td[^>]*>/gi, ': ').replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/\n\s*\n+/g, '\n\n').trim();
}
function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function btn_(href, label, color) {
  return '<a href="' + esc_(href) + '" style="display:inline-block;margin:18px 8px 0 0;padding:11px 20px;background:' + color +
         ';color:#fff;text-decoration:none;border-radius:6px;font-weight:bold">' + esc_(label) + '</a>';
}
function emailHtml_(title, intro, rows, extra) {
  const tr = rows.map(([k, v]) => '<tr><td style="padding:6px 16px 6px 0;color:#5b6780;vertical-align:top">' + esc_(k) +
    '</td><td style="padding:6px 0;color:#16233b">' + esc_(v) + '</td></tr>').join('');
  const head = logoBlob_()
    ? '<span style="display:inline-block;background:#fff;border-radius:6px;padding:6px 10px"><img src="cid:logo" alt="' +
      esc_(HOSPITAL_NAME) + '" style="height:72px;width:auto;display:block"></span>'
    : '<span style="color:#fff;font-weight:bold;font-size:18px">' + esc_(HOSPITAL_NAME) + '</span>';
  return '<div style="font-family:Arial,sans-serif;max-width:560px;color:#16233b">' +
    '<div style="background:' + BRAND_BLUE + ';padding:16px 22px;border-radius:10px 10px 0 0">' + head + '</div>' +
    '<div style="border:1px solid #d5dce8;border-top:0;border-radius:0 0 10px 10px;padding:20px 22px">' +
    '<h2 style="font-family:Georgia,serif;margin:0 0 8px;color:' + BRAND_BLUE + '">' + esc_(title) + '</h2>' +
    '<p style="margin:0 0 16px;color:#5b6780">' + esc_(intro) + '</p>' +
    '<table style="border-collapse:collapse;font-size:14px">' + tr + '</table>' + (extra || '') + '</div></div>';
}
