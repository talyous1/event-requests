/**
 * ============================================================================
 *  VERSION A · Custom web app (WebApp.gs)
 * ============================================================================
 *  Files in this Apps Script project:
 *    Common.gs, WebApp.gs, Index.html, Dashboard.html, Styles.html
 *
 *  Links after you deploy:
 *    Teachers   ->  <web app URL>
 *    Principal  ->  <web app URL>?page=dashboard
 *  (Approvers also see a "Principal dashboard" link on the teacher page.)
 * ============================================================================
 */

/** Run once from the editor. */
function setup() {
  setupTracker_();
  installDigestTrigger_();
  Logger.log('Setup done. Now Deploy > New deployment > Web app (Execute as: Me, Access: Anyone within your domain).');
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  if (page === 'dashboard') return renderPage_('Dashboard', 'Approvals · ' + CONFIG.SCHOOL_NAME);
  return renderPage_('Index', 'Requests · ' + CONFIG.SCHOOL_NAME);
}

/** Link used in teacher emails. */
function teacherLink_(r) {
  const u = webAppUrl_();
  return u ? u + '?page=mine' + (r ? '&id=' + encodeURIComponent(r.id) : '') : '';
}

/** Everything the teacher page needs on load. */
function getTeacherData() {
  const me = currentEmail_();
  const mine = readAll_()
    .filter(r => String(r.email).toLowerCase() === me)
    .sort((a, b) => String(b.submittedAt).localeCompare(String(a.submittedAt)))
    .map(forClient_);
  return {
    email: me,
    name: mine.length ? mine[0].name : '',
    grade: mine.length ? mine[0].grade : '',
    isApprover: isApprover_(me),
    dashboardUrl: dashboardLink_(),
    school: CONFIG.SCHOOL_NAME,
    partyWindow: { start: CONFIG.PARTY_WINDOW_START, end: CONFIG.PARTY_WINDOW_END },
    minDays: CONFIG.MIN_DAYS_NOTICE,
    testMode: CONFIG.TEST_MODE,
    requests: mine,
  };
}

/**
 * TRIAL: run once from the editor to add a few sample requests from made-up
 * teachers, so the Principal dashboard has something to review. No emails are sent.
 * Delete the sample rows from the Requests tab when you're done testing.
 */
function loadSampleRequests() {
  const ymd = d => Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const days = n => { const d = new Date(today); d.setDate(d.getDate() + n); while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1); return d; };
  const thursday = weeksOut => days(((4 - today.getDay() + 7) % 7 || 7) + 7 * weeksOut);
  const now = new Date().toISOString();
  const A = (key, label, value) => ({ key: key, label: label, value: value });

  const samples = [
    { type: 'party', name: 'Ms. Hana (sample)', email: 'sample.teacher1@hqaspring.org', grade: '4th Grade – Ms. Hana',
      title: 'Mystery Day Party', date: ymd(thursday(1)), start: '11:30', end: '12:30', location: 'Classroom',
      answers: [A('coPlanners', 'Co-planning teachers', 'Ms. Noor (4B)'),
        A('activities', 'Learning activities for the day', ['Mystery Day']),
        A('learningPlan', 'Plan for instructional time', 'Logic puzzle stations in the morning, crossword review after recess.'),
        A('food', 'Food, snacks & allergies', 'Fruit cups and popcorn from the class fund. Everything is nut-free.'),
        A('confirm', 'Not shared until approved', 'Confirmed')] },
    { type: 'event', name: 'Mr. Omar (sample)', email: 'sample.teacher2@hqaspring.org', grade: '3rd Grade – Mr. Omar',
      title: 'Classroom Spelling Bee', date: ymd(days(20)), start: '09:00', end: '10:30', location: 'Cafeteria',
      answers: [A('grades', 'Grades / classes involved', '3rd–5th'),
        A('description', 'What’s happening?', 'Class-level spelling bee. Winners move on to the school bee. Held in the cafeteria so testing isn’t interrupted.'),
        A('attendance', 'Expected attendance', '75'),
        A('parents', 'Will parents attend?', 'Livestream only'),
        A('needs', 'What do you need?', ['Tables & chairs', 'Mic / speakers', 'Livestream'])] },
    { type: 'volunteer', name: 'Ms. Sara (sample)', email: 'sample.teacher3@hqaspring.org', grade: '2nd Grade – Ms. Sara',
      title: 'High school student: Yusuf Rahman', date: ymd(days(3)), start: '08:30', end: '11:30', location: 'My classroom only',
      answers: [A('volunteerType', 'Who is coming?', 'High school student'),
        A('volunteerNames', 'Name(s)', 'Yusuf Rahman'),
        A('tasks', 'What will they help with or observe?', 'Help with small reading groups during the morning block.'),
        A('areas', 'Areas they need access to', ['My classroom only']),
        A('confirm', 'Access limits explained', 'Confirmed')] },
    { type: 'fieldtrip', name: 'Ms. Noor (sample)', email: 'sample.teacher4@hqaspring.org', grade: '5th Grade – Ms. Noor',
      title: 'Field Trip: Houston Arboretum', date: ymd(days(30)), start: '08:45', end: '13:30', location: 'Houston Arboretum',
      answers: [A('destination', 'Destination', 'Houston Arboretum'), A('grades', 'Grades / classes', '5th'),
        A('students', 'Number of students', '38'), A('chaperones', 'Chaperones', 'Ms. Noor, Mr. Bilal, 3 parent chaperones'),
        A('transport', 'Transportation', 'School bus'), A('cost', 'Cost per student', '$8'),
        A('purpose', 'Learning purpose', 'Ecosystems unit: field observation journals.')],
      status: STATUS.APPROVED, decidedBy: 'Principal (trial)', note: 'Approved. Please send permission slips home by Friday.' },
  ];

  samples.forEach(s => {
    const r = Object.assign({ id: newId_(), submittedAt: now, status: STATUS.PENDING, note: '', decidedBy: '', decidedAt: '',
      history: [{ at: now, by: s.name, action: 'Submitted' }] }, s);
    if (r.status !== STATUS.PENDING) {
      r.decidedAt = now;
      r.history.push({ at: now, by: r.decidedBy, action: r.status, note: r.note });
    }
    saveRequest_(r);
  });
}

/** New request, or an edit/resubmission of one of the teacher's own requests. */
function submitRequest(data) {
  const me = currentEmail_();
  if (!me) throw new Error('Please sign in with your school Google account.');
  const d = validate_(data);

  let r, resubmitted = false;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (data.id) {
      r = getRequest_(data.id);
      if (!r || String(r.email).toLowerCase() !== me) throw new Error('That request was not found.');
      if (r.status === STATUS.APPROVED || r.status === STATUS.DENIED) {
        throw new Error('This request was already ' + r.status.toLowerCase() + '. Please submit a new request instead.');
      }
      resubmitted = true;
    } else {
      r = { id: newId_(), submittedAt: new Date().toISOString(), email: me, history: [] };
    }
    const now = new Date().toISOString();
    Object.assign(r, d, { status: STATUS.PENDING, note: '', decidedBy: '', decidedAt: '' });
    r.history = (r.history || []).concat([{ at: now, by: d.name, action: resubmitted ? 'Resubmitted' : 'Submitted' }]);
    saveRequest_(r);
  } finally {
    lock.releaseLock();
  }

  notifyNew_(r, resubmitted);
  return forClient_(r);
}

/** Server-side validation (never trust the browser alone). */
function validate_(x) {
  const s = (v, n) => String(v == null ? '' : v).trim().slice(0, n || 300);
  const ymd = v => /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? String(v) : '';
  const d = {
    type: s(x.type, 20),
    name: s(x.name, 120),
    grade: s(x.grade, 120),
    title: s(x.title, 160),
    date: ymd(x.date),
    endDate: ymd(x.endDate),
    start: hm_(x.start),
    end: hm_(x.end),
    location: s(x.location, 160),
    answers: (Array.isArray(x.answers) ? x.answers : []).slice(0, 40).map(a => ({
      key: s(a.key, 40),
      label: s(a.label, 120),
      value: Array.isArray(a.value) ? a.value.map(v => s(v, 200)).filter(String) : s(a.value, 3000),
    })),
  };
  if (!TYPES[d.type]) throw new Error('Please choose a request type.');
  if (!d.name || !d.grade) throw new Error('Please add your name and class.');
  if (!d.title) throw new Error('Please add a title.');
  if (!d.date) throw new Error('Please choose a date.');
  if (d.endDate && d.endDate < d.date) throw new Error('The end date is before the start date.');
  if (d.start && d.end && !d.endDate && d.end <= d.start) throw new Error('The end time must be after the start time.');

  if (d.type === 'party') {
    if (ymdToDate_(d.date).getDay() !== 4) throw new Error('Class parties can only be held on Thursdays.');
    if (CONFIG.PARTY_WINDOW_START && CONFIG.PARTY_WINDOW_END &&
        (d.start < CONFIG.PARTY_WINDOW_START || d.end > CONFIG.PARTY_WINDOW_END)) {
      throw new Error('Parties must start and end within lunch & recess (' +
        fmtTime_(CONFIG.PARTY_WINDOW_START) + ' – ' + fmtTime_(CONFIG.PARTY_WINDOW_END) + ').');
    }
    const a = k => (d.answers.find(z => z.key === k) || {}).value;
    if (!a('coPlanners')) throw new Error('Parties are planned as a team. Please list your co-planning teachers.');
    if (!a('activities') || !a('activities').length) throw new Error('Please choose the learning activities for the day.');
  }
  return d;
}
