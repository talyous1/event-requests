/**
 * ============================================================================
 *  HQA Spring · Event & Volunteer Requests  —  SHARED BACKEND (Common.gs)
 * ============================================================================
 *  Teachers request approval for class parties, school events, field trips,
 *  and parent/student volunteers or observers. The Principal approves, asks
 *  for changes, or denies from a dashboard. Everything is logged to the
 *  "Requests" tab of the Google Sheet this script is attached to.
 *
 *  1. Edit CONFIG below (emails especially).
 *  2. Run setup() once from the editor and accept the permissions.
 *  3. Deploy > New deployment > Web app
 *       Execute as: Me   ·   Who has access: Anyone within hqaspring.org
 *  4. If you redeploy later, use "Manage deployments" > edit > New version
 *     so the link stays the same.
 * ============================================================================
 */

const CONFIG = {
  // TRIAL MODE: every email goes ONLY to TEST_EMAIL (with a note showing who it
  // would have gone to), and TEST_EMAIL can approve requests. Nobody else is
  // emailed. Set TEST_MODE to false when you're ready to go live.
  TEST_MODE: true,
  TEST_EMAIL: 'finance@hqaspring.org',

  SCHOOL_NAME: 'HQA Spring',
  APP_NAME: 'Event & Volunteer Requests',
  TIMEZONE: 'America/Chicago',

  // Who can approve. Email -> name shown to teachers ("Approved by ...").
  APPROVERS: {
    'principal@hqaspring.org': 'Principal',
  },

  // Get an FYI copy of every new request and every decision.
  ADMIN_CC: ['hr@hqaspring.org', 'admin@hqaspring.org', 'secretary@hqaspring.org'],

  // Approved volunteer/observer visits go to the Front Desk, Principal and Ms. Omayya.
  VOLUNTEER_NOTIFY: ['frontdesk@hqaspring.org', 'principal@hqaspring.org', 'omayya@hqaspring.org'],

  // Optional: a shared Google Calendar ID. Approved requests are added to it
  // automatically (and removed if the decision changes). Leave '' to skip.
  CALENDAR_ID: '',

  // Class party window (lunch + recess), 24-hour "HH:MM". Leave '' to skip the check.
  PARTY_WINDOW_START: '11:30',
  PARTY_WINDOW_END: '12:45',

  // Days of notice you'd like. Shown as a check on each request (not a hard block).
  MIN_DAYS_NOTICE: 7,

  // Weekday-morning email to approvers listing anything still pending.
  DAILY_DIGEST: true,
  DIGEST_HOUR: 7,

  // Normally detected automatically. Paste the /exec link here if emails show no button.
  WEB_APP_URL: '',
};

/* ---------------------------------------------------------------------------
 *  Constants
 * ------------------------------------------------------------------------- */
const STATUS = {
  PENDING: 'Pending',
  APPROVED: 'Approved',
  CHANGES: 'Needs Changes',
  DENIED: 'Denied',
};

const TYPES = {
  party:     { label: 'Class Party',          color: '#B45309' },
  event:     { label: 'School Event',         color: '#1D4ED8' },
  fieldtrip: { label: 'Field Trip',           color: '#6D28D9' },
  volunteer: { label: 'Volunteer / Observer', color: '#0F766E' },
};

const STATUS_STYLE = {
  'Pending':       { bg: '#FEF3C7', fg: '#92400E' },
  'Approved':      { bg: '#DCFCE7', fg: '#166534' },
  'Needs Changes': { bg: '#FFEDD5', fg: '#9A3412' },
  'Denied':        { bg: '#FEE2E2', fg: '#991B1B' },
};

const SHEET_NAME = 'Requests';
const HEADERS = [
  'Request ID', 'Submitted', 'Status', 'Type', 'Title', 'Requested By', 'Email',
  'Class / Grade', 'Date', 'Time', 'Location', 'Summary',
  'Decided By', 'Decided At', 'Principal Note', 'Data',
];
const COL = HEADERS.reduce((m, h, i) => (m[h] = i + 1, m), {});

/* ---------------------------------------------------------------------------
 *  Sheet setup
 * ------------------------------------------------------------------------- */
function setupTracker_() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  const rows = sh.getMaxRows() - 1;

  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontFamily('Montserrat').setFontWeight('bold').setFontSize(10)
    .setFontColor('#E9F6E6').setBackground('#9EAF9A').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 34);
  sh.setHiddenGridlines(true);
  sh.setTabColor('#9EAF9A');
  sh.getRange(2, 1, rows, HEADERS.length).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('top');

  const widths = { 'Request ID': 120, 'Submitted': 150, 'Status': 115, 'Type': 150, 'Title': 220,
    'Requested By': 140, 'Email': 190, 'Class / Grade': 140, 'Date': 130, 'Time': 150,
    'Location': 130, 'Summary': 340, 'Decided By': 120, 'Decided At': 150, 'Principal Note': 260 };
  Object.keys(widths).forEach(h => sh.setColumnWidth(COL[h], widths[h]));
  sh.getRange(2, COL['Summary'], rows, 1).setWrap(true);
  sh.getRange(2, COL['Principal Note'], rows, 1).setWrap(true);
  sh.getRange(2, COL['Submitted'], rows, 1).setNumberFormat('mmm d, yyyy h:mm AM/PM');
  sh.getRange(2, COL['Decided At'], rows, 1).setNumberFormat('mmm d, yyyy h:mm AM/PM');
  sh.getRange(2, COL['Date'], rows, 1).setNumberFormat('ddd, mmm d, yyyy');
  sh.hideColumns(COL['Data']);

  // Thin pale-green row lines
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length - 1)
    .setBorder(null, null, null, null, null, true, '#DCE8D8', SpreadsheetApp.BorderStyle.SOLID);

  // Status dropdown + colors
  const statusRange = sh.getRange(2, COL['Status'], rows, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(STATUS), true).setAllowInvalid(false).build());
  sh.setConditionalFormatRules(Object.keys(STATUS_STYLE).map(s =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s)
      .setBackground(STATUS_STYLE[s].bg).setFontColor(STATUS_STYLE[s].fg)
      .setRanges([statusRange]).build()));
  return sh;
}

function installDigestTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendPendingDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (CONFIG.DAILY_DIGEST) {
    ScriptApp.newTrigger('sendPendingDigest').timeBased().everyDays(1)
      .atHour(CONFIG.DIGEST_HOUR).inTimezone(CONFIG.TIMEZONE).create();
  }
}

/* ---------------------------------------------------------------------------
 *  Storage
 * ------------------------------------------------------------------------- */
function getSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME) || setupTracker_();
}

function newId_() {
  const stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMM');
  return 'REQ-' + stamp + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
}

function findRow_(id) {
  const sh = getSheet_();
  if (!id || sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, COL['Request ID'], sh.getLastRow() - 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function rowValues_(r) {
  return [
    r.id,
    new Date(r.submittedAt),
    r.status,
    TYPES[r.type].label,
    r.title,
    r.name,
    r.email,
    r.grade,
    r.date ? ymdToDate_(r.date) : '',
    timeLabel_(r),
    r.location || '',
    summaryText_(r),
    r.decidedBy || '',
    r.decidedAt ? new Date(r.decidedAt) : '',
    r.note || '',
    JSON.stringify(r),
  ];
}

function saveRequest_(r) {
  const sh = getSheet_();
  const row = findRow_(r.id);
  if (row) sh.getRange(row, 1, 1, HEADERS.length).setValues([rowValues_(r)]);
  else sh.appendRow(rowValues_(r));
}

/** The sheet's Status / Decided / Note columns win, so manual edits in the sheet are respected. */
function parseRow_(v) {
  let r;
  try { r = JSON.parse(v[COL['Data'] - 1]); } catch (e) { return null; }
  if (!r || !r.id) return null;
  r.status = v[COL['Status'] - 1] || STATUS.PENDING;
  r.note = v[COL['Principal Note'] - 1] || '';
  r.decidedBy = v[COL['Decided By'] - 1] || '';
  const at = v[COL['Decided At'] - 1];
  r.decidedAt = at instanceof Date ? at.toISOString() : (r.decidedAt || '');
  return r;
}

function readAll_() {
  const sh = getSheet_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues()
    .map(parseRow_).filter(Boolean);
}

function getRequest_(id) {
  const row = findRow_(id);
  if (!row) return null;
  return parseRow_(getSheet_().getRange(row, 1, 1, HEADERS.length).getValues()[0]);
}

/* ---------------------------------------------------------------------------
 *  Formatting helpers
 * ------------------------------------------------------------------------- */
function ymdToDate_(ymd) {
  const p = String(ymd).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function hm_(t) {
  const m = String(t || '').match(/(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

function fmtTime_(t) {
  if (!t) return '';
  const p = t.split(':').map(Number);
  const h = p[0] % 12 || 12;
  return h + ':' + ('0' + p[1]).slice(-2) + (p[0] < 12 ? ' AM' : ' PM');
}

function fmtDate_(ymd, pattern) {
  return ymd ? Utilities.formatDate(ymdToDate_(ymd), CONFIG.TIMEZONE, pattern || 'EEE, MMM d, yyyy') : '';
}

function dateLabel_(r) {
  if (!r.date) return '';
  if (r.endDate && r.endDate !== r.date) return fmtDate_(r.date, 'EEE, MMM d') + ' – ' + fmtDate_(r.endDate);
  return fmtDate_(r.date, 'EEEE, MMM d, yyyy');
}

function timeLabel_(r) {
  if (r.start && r.end) return fmtTime_(r.start) + ' – ' + fmtTime_(r.end);
  return fmtTime_(r.start || r.end || '');
}

function ans_(r, key) {
  const a = (r.answers || []).find(x => x.key === key);
  if (!a) return '';
  return Array.isArray(a.value) ? a.value.join(', ') : String(a.value || '');
}

function answerText_(a) {
  return Array.isArray(a.value) ? a.value.join(', ') : String(a.value == null ? '' : a.value);
}

function summaryText_(r) {
  return (r.answers || []).filter(a => answerText_(a))
    .map(a => a.label + ': ' + answerText_(a)).join('\n');
}

function detailRows_(r) {
  return [
    ['Request', TYPES[r.type].label + ' · ' + r.id],
    ['Requested by', r.name + (r.email ? ' (' + r.email + ')' : '')],
    ['Class / grade', r.grade],
    ['Date', dateLabel_(r)],
    ['Time', timeLabel_(r)],
    ['Location', r.location],
  ].concat((r.answers || []).map(a => [a.label, answerText_(a)]));
}

function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function daysNotice_(r) {
  const sub = new Date(r.submittedAt);
  const subDay = new Date(sub.getFullYear(), sub.getMonth(), sub.getDate());
  return Math.round((ymdToDate_(r.date) - subDay) / 864e5);
}

/* ---------------------------------------------------------------------------
 *  Policy checks (from the Nov 6 staff meeting)
 * ------------------------------------------------------------------------- */
function policyChecks_(r) {
  const checks = [];
  if (!r.date) return checks;

  if (r.type === 'party') {
    const day = Utilities.formatDate(ymdToDate_(r.date), CONFIG.TIMEZONE, 'EEEE');
    checks.push({ ok: day === 'Thursday', label: 'Held on a Thursday',
      detail: day === 'Thursday' ? '' : 'This party is on a ' + day + '.' });

    if (CONFIG.PARTY_WINDOW_START && CONFIG.PARTY_WINDOW_END && r.start && r.end) {
      const inside = r.start >= CONFIG.PARTY_WINDOW_START && r.end <= CONFIG.PARTY_WINDOW_END;
      checks.push({ ok: inside, label: 'Within lunch & recess',
        detail: inside ? '' : 'Lunch & recess is ' + fmtTime_(CONFIG.PARTY_WINDOW_START) + ' – ' + fmtTime_(CONFIG.PARTY_WINDOW_END) + '.' });
    }
    const team = ans_(r, 'coPlanners');
    checks.push({ ok: !!team, label: 'Planned as a team', detail: team ? 'With ' + team : 'No co-planning teachers listed.' });
    const acts = ans_(r, 'activities');
    checks.push({ ok: !!acts && !!ans_(r, 'learningPlan'), label: 'Learning plan for the day',
      detail: acts || 'No learning activities listed.' });
  }

  if (r.type === 'fieldtrip') {
    const ch = ans_(r, 'chaperones');
    checks.push({ ok: !!ch, label: 'Chaperones listed', detail: ch ? '' : 'No chaperones listed.' });
  }

  const days = daysNotice_(r);
  checks.push({ ok: days >= CONFIG.MIN_DAYS_NOTICE,
    label: CONFIG.MIN_DAYS_NOTICE + '+ days notice',
    detail: days < 0 ? 'The date had already passed when this was submitted.'
      : days + (days === 1 ? ' day' : ' days') + ' before the date.' });
  return checks;
}

/* ---------------------------------------------------------------------------
 *  People & links
 * ------------------------------------------------------------------------- */
function currentEmail_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase();
}

function approverEmails_() {
  const list = Object.keys(CONFIG.APPROVERS);
  return uniq_(CONFIG.TEST_MODE ? list.concat([CONFIG.TEST_EMAIL]) : list);
}

function isApprover_(email) {
  return !!email && approverEmails_().indexOf(String(email).toLowerCase()) > -1;
}

function approverName_(email) {
  const key = Object.keys(CONFIG.APPROVERS).find(e => e.toLowerCase() === String(email).toLowerCase());
  if (key) return CONFIG.APPROVERS[key];
  if (CONFIG.TEST_MODE && String(email).toLowerCase() === String(CONFIG.TEST_EMAIL).toLowerCase()) return 'Principal (trial)';
  return email;
}

function webAppUrl_() {
  if (CONFIG.WEB_APP_URL) return CONFIG.WEB_APP_URL;
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function dashboardLink_(r) {
  const u = webAppUrl_();
  return u ? u + '?page=dashboard' + (r ? '&id=' + encodeURIComponent(r.id) : '') : '';
}

function uniq_(list) {
  const seen = {};
  return list.map(e => String(e || '').trim().toLowerCase())
    .filter(e => e && !seen[e] && (seen[e] = true));
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function renderPage_(file, title) {
  return HtmlService.createTemplateFromFile(file).evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Shape a request for the browser (adds labels and checks). */
function forClient_(r) {
  return Object.assign({}, r, {
    typeLabel: TYPES[r.type].label,
    typeColor: TYPES[r.type].color,
    dateLabel: dateLabel_(r),
    shortDate: fmtDate_(r.date, 'EEE, MMM d'),
    timeLabel: timeLabel_(r),
    submittedLabel: Utilities.formatDate(new Date(r.submittedAt), CONFIG.TIMEZONE, 'MMM d, h:mm a'),
    decidedLabel: r.decidedAt ? Utilities.formatDate(new Date(r.decidedAt), CONFIG.TIMEZONE, 'MMM d, h:mm a') : '',
    checks: policyChecks_(r),
  });
}

/* ---------------------------------------------------------------------------
 *  Principal dashboard API
 * ------------------------------------------------------------------------- */
function getDashboardData() {
  const me = currentEmail_();
  if (!isApprover_(me)) return { denied: true, email: me };
  const list = readAll_().map(forClient_)
    .sort((a, b) => String(b.submittedAt).localeCompare(String(a.submittedAt)));
  return { email: me, name: approverName_(me), school: CONFIG.SCHOOL_NAME, testMode: CONFIG.TEST_MODE, requests: list };
}

/**
 * Approve / ask for changes / deny.
 * Returns {conflict:true} if someone already decided, unless force is true.
 */
function decide(id, decision, note, force) {
  const me = currentEmail_();
  if (!isApprover_(me)) throw new Error('Only approvers can make decisions.');
  if ([STATUS.APPROVED, STATUS.CHANGES, STATUS.DENIED].indexOf(decision) < 0) throw new Error('Unknown decision.');
  note = String(note || '').trim().slice(0, 2000);
  if (decision === STATUS.CHANGES && !note) throw new Error('Add a note so the teacher knows what to change.');

  let r;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    r = getRequest_(id);
    if (!r) throw new Error('Request ' + id + ' was not found.');
    if (r.status !== STATUS.PENDING && !force) {
      return { conflict: true, status: r.status, decidedBy: r.decidedBy };
    }
    r.status = decision;
    r.note = note;
    r.decidedBy = approverName_(me);
    r.decidedByEmail = me;
    r.decidedAt = new Date().toISOString();
    r.history = (r.history || []).concat([{ at: r.decidedAt, by: r.decidedBy, action: decision, note: note }]);
    saveRequest_(r);
  } finally {
    lock.releaseLock();
  }

  syncCalendar_(r);
  notifyDecision_(r);
  return { ok: true, request: forClient_(r) };
}

/* ---------------------------------------------------------------------------
 *  Calendar (optional)
 * ------------------------------------------------------------------------- */
function syncCalendar_(r) {
  if (!CONFIG.CALENDAR_ID) return;
  try {
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    if (!cal) return;
    if (r.status === STATUS.APPROVED && !r.calendarEventId) {
      const title = TYPES[r.type].label + ': ' + r.title;
      const opts = { location: r.location || '', description:
        'Requested by ' + r.name + ' (' + r.grade + ')\nApproved by ' + r.decidedBy + '\n\n' + summaryText_(r) };
      const at = (ymd, t) => { const d = ymdToDate_(ymd); const p = t.split(':'); d.setHours(+p[0], +p[1]); return d; };
      const ev = (r.start && r.end)
        ? cal.createEvent(title, at(r.date, r.start), at(r.endDate || r.date, r.end), opts)
        : cal.createAllDayEvent(title, ymdToDate_(r.date), opts);
      r.calendarEventId = ev.getId();
    } else if (r.status !== STATUS.APPROVED && r.calendarEventId) {
      const ev = cal.getEventById(r.calendarEventId);
      if (ev) ev.deleteEvent();
      r.calendarEventId = '';
    } else {
      return;
    }
    saveRequest_(r);
  } catch (err) {
    console.warn('Calendar sync failed for ' + r.id + ': ' + err);
  }
}

/* ---------------------------------------------------------------------------
 *  Emails
 * ------------------------------------------------------------------------- */
function emailHtml_(o) {
  const esc = escapeHtml_;
  const accent = o.accent || '#1E459D';
  const rows = (o.rows || []).filter(r => r[1] !== '' && r[1] != null).map(r =>
    '<tr><td style="padding:9px 14px;border-bottom:1px solid #EBEFF6;color:#5B667A;font-size:13px;width:36%;vertical-align:top;">' + esc(r[0]) +
    '</td><td style="padding:9px 14px;border-bottom:1px solid #EBEFF6;color:#172033;font-size:14px;vertical-align:top;">' + esc(r[1]).replace(/\n/g, '<br>') + '</td></tr>').join('');

  const checks = (o.checks || []).map(c =>
    '<tr><td style="padding:5px 0;vertical-align:top;width:26px;"><span style="display:inline-block;width:18px;height:18px;line-height:18px;border-radius:9px;text-align:center;font-size:12px;font-weight:bold;color:#fff;background-color:' +
    (c.ok ? '#16A34A' : '#EA580C') + ';">' + (c.ok ? '&#10003;' : '!') + '</span></td><td style="padding:5px 0;font-size:14px;color:#172033;">' +
    '<b>' + esc(c.label) + '</b>' + (c.detail ? '<span style="color:#5B667A;"> · ' + esc(c.detail) + '</span>' : '') + '</td></tr>').join('');

  const status = o.status ? '<span style="display:inline-block;padding:4px 12px;border-radius:999px;font-size:12px;font-weight:bold;background-color:' +
    STATUS_STYLE[o.status].bg + ';color:' + STATUS_STYLE[o.status].fg + ';">' + esc(o.status) + '</span>' : '';

  const note = o.note ? '<div style="margin:18px 0 0;padding:14px 16px;border-left:4px solid ' + accent +
    ';background-color:#F4F6FB;border-radius:6px;"><div style="font-size:12px;color:#5B667A;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;">' +
    esc(o.noteLabel || 'Note') + '</div><div style="font-size:15px;color:#172033;">' + esc(o.note).replace(/\n/g, '<br>') + '</div></div>' : '';

  const button = (o.button && o.button.url) ? '<div style="margin:24px 0 4px;"><a href="' + esc(o.button.url) +
    '" style="display:inline-block;background-color:' + accent + ';color:#ffffff;text-decoration:none;font-weight:bold;font-size:15px;padding:12px 26px;border-radius:8px;">' +
    esc(o.button.label) + '</a></div>' : '';

  return '<!doctype html><html><body bgcolor="#EEF2F8" style="margin:0;padding:0;background-color:#EEF2F8;">' +
    '<table role="presentation" width="100%" bgcolor="#EEF2F8" cellpadding="0" cellspacing="0" style="background-color:#EEF2F8;padding:24px 12px;"><tr><td align="center">' +
    '<table role="presentation" width="100%" bgcolor="#FFFFFF" cellpadding="0" cellspacing="0" style="max-width:600px;background-color:#FFFFFF;border-radius:14px;overflow:hidden;font-family:Arial,Helvetica,sans-serif;">' +
    '<tr><td bgcolor="#17377F" align="center" style="background-color:#17377F;padding:30px 26px 24px;text-align:center;">' +
    '<img src="cid:hqalogo" width="72" height="72" alt="HQA Spring" style="display:block;margin:0 auto 12px;border:0;width:72px;height:72px;">' +
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#FFFFFF;font-size:19px;font-weight:bold;letter-spacing:.02em;">Houston Quran Academy Spring</div>' +
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#B9C7EA;font-size:12px;letter-spacing:.16em;text-transform:uppercase;margin-top:6px;">' + esc(CONFIG.APP_NAME) + '</div></td></tr>' +
    '<tr><td bgcolor="#E03A2C" style="background-color:#E03A2C;height:5px;font-size:0;line-height:0;">&nbsp;</td></tr>' +
    '<tr><td style="padding:26px 26px 8px;">' +
    (o.eyebrow ? '<div style="font-size:12px;color:' + accent + ';font-weight:bold;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;">' + esc(o.eyebrow) + '</div>' : '') +
    '<div style="font-size:22px;line-height:1.3;font-weight:bold;color:#172033;margin:0 0 8px;">' + esc(o.heading) + '</div>' +
    status +
    (o.intro ? '<p style="font-size:15px;line-height:1.55;color:#34405A;margin:14px 0 0;">' + o.intro + '</p>' : '') +
    note + button + '</td></tr>' +
    (checks ? '<tr><td style="padding:16px 26px 0;"><div style="font-size:12px;color:#5B667A;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Policy checks</div><table role="presentation" cellpadding="0" cellspacing="0">' + checks + '</table></td></tr>' : '') +
    (rows ? '<tr><td style="padding:18px 26px 6px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #DFE5F0;border-radius:10px;border-collapse:separate;overflow:hidden;">' + rows + '</table></td></tr>' : '') +
    '<tr><td style="padding:18px 26px 26px;font-size:12px;color:#8792A8;">' + (o.footer || 'Automated message from the ' + esc(CONFIG.SCHOOL_NAME) + ' request system.') + '</td></tr>' +
    '</table></td></tr></table></body></html>';
}

function send_(to, subject, html, cc) {
  to = uniq_([].concat(to || []));
  cc = uniq_([].concat(cc || [])).filter(e => to.indexOf(e) < 0);
  if (!to.length) return;
  if (CONFIG.TEST_MODE) {
    const banner = '<div style="background-color:#FEF3C7;color:#92400E;font:13px Arial,sans-serif;padding:10px 16px;text-align:center;">' +
      '<b>TRIAL MODE</b> · Only you received this. It would have gone to: ' + escapeHtml_(to.join(', ')) +
      (cc.length ? ' · cc: ' + escapeHtml_(cc.join(', ')) : '') + '</div>';
    html = html.replace('<body bgcolor="#EEF2F8" style="margin:0;padding:0;background-color:#EEF2F8;">', '$&' + banner);
    subject = '[TRIAL] ' + subject;
    to = [String(CONFIG.TEST_EMAIL).toLowerCase()];
    cc = [];
  }
  const msg = { to: to.join(','), subject: subject, htmlBody: html, name: CONFIG.SCHOOL_NAME + ' Requests' };
  if (cc.length) msg.cc = cc.join(',');
  if (html.indexOf('cid:hqalogo') > -1) msg.inlineImages = { hqalogo: logoBlob_() };
  MailApp.sendEmail(msg);
}

function subjectTail_(r) {
  return r.title + ' (' + fmtDate_(r.date, 'EEE, MMM d') + ')';
}

/** New or resubmitted request: approvers get a review link, the teacher gets a receipt. */
function notifyNew_(r, resubmitted) {
  const checks = policyChecks_(r);
  const flags = checks.filter(c => !c.ok).length;
  const accent = TYPES[r.type].color;

  send_(approverEmails_(),
    (resubmitted ? '[Resubmitted] ' : '[Needs approval] ') + TYPES[r.type].label + ': ' + subjectTail_(r),
    emailHtml_({
      accent: accent,
      eyebrow: TYPES[r.type].label + (resubmitted ? ' · resubmitted' : ''),
      heading: r.title,
      status: STATUS.PENDING,
      intro: escapeHtml_(r.name) + ' (' + escapeHtml_(r.grade) + ') ' + (resubmitted ? 'updated and resubmitted' : 'submitted') +
        ' a request for <b>' + escapeHtml_(dateLabel_(r)) + '</b>.' +
        (flags ? ' <span style="color:#C2410C;font-weight:bold;">' + flags + ' policy check' + (flags > 1 ? 's need' : ' needs') + ' a look.</span>' : ''),
      checks: checks,
      rows: detailRows_(r),
      button: { label: 'Review request', url: dashboardLink_(r) },
    }),
    CONFIG.ADMIN_CC);

  const link = teacherLink_(r);
  send_(r.email, 'Request received: ' + subjectTail_(r), emailHtml_({
    accent: '#1E459D',
    eyebrow: 'Request received',
    heading: r.title,
    status: STATUS.PENDING,
    intro: 'Assalamu alaikum ' + escapeHtml_(greetName_(r.name)) + ', your request was sent to the Principal. ' +
      'You will get an email as soon as it is reviewed. Please don\'t share it with students or families until it\'s approved.',
    checks: flags ? checks.filter(c => !c.ok) : null,
    rows: detailRows_(r),
    button: link ? { label: 'View my requests', url: link } : null,
  }));
}

function notifyDecision_(r) {
  const word = { 'Approved': 'approved', 'Needs Changes': 'sent back for changes', 'Denied': 'not approved' }[r.status];
  const accent = { 'Approved': '#16A34A', 'Needs Changes': '#EA580C', 'Denied': '#DC2626' }[r.status];
  const link = teacherLink_(r);

  let intro = 'Your request was <b>' + word + '</b> by ' + escapeHtml_(r.decidedBy) + '.';
  if (r.status === STATUS.APPROVED && r.type === 'party') intro += ' You can now share the party with students and families.';
  if (r.status === STATUS.APPROVED && r.type === 'volunteer') intro += ' The Front Desk has been notified. Please let your volunteer know their access to the school is limited to what they\'re helping with' +
    (/student/i.test(ans_(r, 'volunteerType')) ? ', and student volunteers may not have their phones during school hours.' : '.');
  if (r.status === STATUS.CHANGES) intro += ' Please update it using the note below and resubmit.';

  send_(r.email, r.status + ': ' + subjectTail_(r), emailHtml_({
    accent: accent,
    eyebrow: TYPES[r.type].label,
    heading: r.title,
    status: r.status,
    intro: intro,
    note: r.note,
    noteLabel: 'Note from ' + r.decidedBy,
    rows: detailRows_(r),
    button: link ? { label: r.status === STATUS.CHANGES ? 'Edit & resubmit' : 'View my requests', url: link } : null,
  }));

  const others = approverEmails_().filter(e => e !== String(r.decidedByEmail || '').toLowerCase());
  send_(CONFIG.ADMIN_CC.concat(others), r.status + ' by ' + r.decidedBy + ': ' + subjectTail_(r), emailHtml_({
    accent: accent,
    eyebrow: 'FYI · ' + TYPES[r.type].label,
    heading: r.title,
    status: r.status,
    intro: escapeHtml_(r.decidedBy) + ' marked ' + escapeHtml_(r.name) + '\'s request as <b>' + escapeHtml_(r.status) + '</b>.',
    note: r.note,
    noteLabel: 'Principal note',
    rows: detailRows_(r),
    button: { label: 'Open dashboard', url: dashboardLink_(r) },
  }));

  if (r.status === STATUS.APPROVED && r.type === 'volunteer' && CONFIG.VOLUNTEER_NOTIFY.length) {
    send_(CONFIG.VOLUNTEER_NOTIFY, 'Approved visitor: ' + ans_(r, 'volunteerNames') + ' · ' + fmtDate_(r.date, 'EEE, MMM d'), emailHtml_({
      accent: TYPES.volunteer.color,
      eyebrow: 'Approved volunteer / observer',
      heading: ans_(r, 'volunteerNames') || r.title,
      status: STATUS.APPROVED,
      intro: 'Please expect this visitor at the front desk. They are helping <b>' + escapeHtml_(r.name) + '</b> (' + escapeHtml_(r.grade) +
        ') and their access is limited to the areas listed below.',
      note: r.note, noteLabel: 'Principal note',
      rows: detailRows_(r),
    }));
  }
}

/** Weekday-morning reminder of anything still waiting. */
function sendPendingDigest() {
  const dow = Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'u')); // 1 = Mon … 7 = Sun
  if (dow > 5) return;
  const pending = readAll_().filter(r => r.status === STATUS.PENDING)
    .sort((a, b) => String(a.date).localeCompare(String(b.date)));
  if (!pending.length) return;
  send_(approverEmails_(), pending.length + ' request' + (pending.length > 1 ? 's' : '') + ' waiting for approval', emailHtml_({
    eyebrow: 'Morning summary',
    heading: pending.length + ' request' + (pending.length > 1 ? 's are' : ' is') + ' waiting for you',
    intro: 'Sorted by event date, soonest first.',
    rows: pending.map(r => [fmtDate_(r.date, 'EEE, MMM d'), TYPES[r.type].label + ': ' + r.title + '\n' + r.name +
      (policyChecks_(r).some(c => !c.ok) ? '  ·  has policy flags' : '')]),
    button: { label: 'Open dashboard', url: dashboardLink_() },
  }));
}

/** "Ms. Hana Ali" -> "Ms. Hana", "Hana Ali" -> "Hana" (for greetings). */
function greetName_(name) {
  const w = String(name || '').trim().split(/\s+/);
  return /^(ms|mr|mrs|miss|dr|sr|br|sh|ustadh|ustadha|sister|brother)\.?$/i.test(w[0]) && w[1] ? w[0] + ' ' + w[1] : w[0];
}
