/**
 * ============================================================================
 *  VERSION B · Google Form + script (GoogleForm.gs)
 * ============================================================================
 *  Files in this Apps Script project:
 *    Common.gs, GoogleForm.gs, Dashboard.html, Styles.html
 *
 *  setup() builds the whole Google Form for you (with a separate page for
 *  each request type), links it to this spreadsheet, and turns on emails.
 *  The form link and edit link are printed in the execution log.
 *
 *  Links after you deploy the web app:
 *    Teachers   ->  the Google Form link
 *    Principal  ->  <web app URL>?page=dashboard   (or just the web app URL)
 *
 *  Google Forms can't block a date by weekday, so non-Thursday parties and
 *  times outside lunch/recess are flagged for the Principal instead.
 * ============================================================================
 */

const Q = {
  name: 'Your name',
  grade: 'Class / grade',
  type: 'What do you need approved?',
  // Class party
  pTitle: 'Party name (optional)', pDate: 'Party date', pStart: 'Party starts', pEnd: 'Party ends',
  pLoc: 'Party location', pTeam: 'Co-planning teachers', pActs: 'Learning activities for the day',
  pPlan: 'Plan for instructional time', pFood: 'Food, snacks & allergies', pParents: 'Parent help',
  pConfirm: 'Not shared until approved',
  // School event
  eTitle: 'Event name', eDate: 'Event date', eStart: 'Event starts', eEnd: 'Event ends', eLoc: 'Event location',
  eGrades: 'Grades / classes involved', eDesc: 'What’s happening?', eAtt: 'Expected attendance',
  eBudget: 'Estimated cost', eParents: 'Will parents attend?', eNeeds: 'What do you need?',
  // Field trip
  tDest: 'Destination', tAddr: 'Address', tDate: 'Trip date', tStart: 'Leave school', tEnd: 'Back at school',
  tGrades: 'Grades / classes on the trip', tStudents: 'Number of students', tChap: 'Chaperones',
  tTrans: 'Transportation', tCost: 'Cost per student', tSlips: 'Permission slips', tPurpose: 'Learning purpose',
  // Volunteer / observer
  vType: 'Who is coming?', vNames: 'Volunteer name(s)', vContact: 'Volunteer phone or email',
  vDate: 'First day', vEnd: 'Last day (only if more than one day)', vStart: 'Arrives', vLeave: 'Leaves',
  vTasks: 'What will they help with or observe?', vAreas: 'Areas they need access to', vConfirm: 'Access limits explained',
  // Last page
  notes: 'Anything else the Principal should know?',
  resubmit: 'Request ID (only when resubmitting)',
};

const TYPE_CHOICES = {
  'Class Party': 'party',
  'School Event': 'event',
  'Field Trip': 'fieldtrip',
  'Volunteer / Observer': 'volunteer',
};

/** Run once from the editor. */
function setup() {
  setupTracker_();
  installDigestTrigger_();
  const form = getForm_() || createForm_();
  Logger.log('Teacher form link:  ' + form.getPublishedUrl());
  Logger.log('Edit the form here: ' + form.getEditUrl());
  Logger.log('Next: Deploy > New deployment > Web app (Execute as: Me, Access: Anyone within your domain) for the Principal dashboard.');
}

function doGet() {
  return renderPage_('Dashboard', 'Approvals · ' + CONFIG.SCHOOL_NAME);
}

function getForm_() {
  const id = PropertiesService.getScriptProperties().getProperty('FORM_ID');
  if (!id) return null;
  try { return FormApp.openById(id); } catch (e) { return null; }
}

/* ---------------------------------------------------------------------------
 *  Build the form
 * ------------------------------------------------------------------------- */
function createForm_() {
  const winText = CONFIG.PARTY_WINDOW_START
    ? ' (' + fmtTime_(CONFIG.PARTY_WINDOW_START) + ' – ' + fmtTime_(CONFIG.PARTY_WINDOW_END) + ')' : '';

  const form = FormApp.create(CONFIG.SCHOOL_NAME + ' · Event & Volunteer Request');
  form.setDescription('Class parties, school events, field trips and volunteers all need the Principal’s approval ' +
      'before you share them with students or families. You’ll get an email as soon as your request is reviewed.')
    .setCollectEmail(true)
    .setProgressBar(true)
    .setShowLinkToRespondAgain(true)
    .setConfirmationMessage('Thank you! Your request was sent to the Principal. Check your email for a copy, and for the decision once it’s reviewed.');
  try { form.setRequireLogin(true); } catch (e) { /* Workspace-only setting */ }

  const text = (title, help, req) => form.addTextItem().setTitle(title).setHelpText(help || '').setRequired(!!req);
  const para = (title, help, req) => form.addParagraphTextItem().setTitle(title).setHelpText(help || '').setRequired(!!req);
  const date = (title, help, req) => form.addDateItem().setTitle(title).setHelpText(help || '').setRequired(!!req);
  const time = (title, help, req) => form.addTimeItem().setTitle(title).setHelpText(help || '').setRequired(!!req);
  const one = (title, opts, req, help) => form.addMultipleChoiceItem().setTitle(title).setChoiceValues(opts).setHelpText(help || '').setRequired(!!req);
  const many = (title, opts, req, help, other) => form.addCheckboxItem().setTitle(title).setChoiceValues(opts).setHelpText(help || '').setRequired(!!req).showOtherOption(!!other);
  const number = (title, req) => form.addTextItem().setTitle(title).setRequired(!!req)
    .setValidation(FormApp.createTextValidation().requireNumber().setHelpText('Numbers only').build());
  const LOCS = ['Classroom', 'Cafeteria', 'Gym', 'Playground', 'Library', 'Prayer hall'];

  // Page 1 · About you
  text(Q.name, '', true);
  text(Q.grade, 'e.g. 3rd Grade – Ms. Sara', true);
  const typeItem = form.addMultipleChoiceItem().setTitle(Q.type).setRequired(true);

  // Page 2 · Class party
  const pParty = form.addPageBreakItem().setTitle('Class Party').setHelpText(
    'Party rules (Nov 6 staff meeting):\n' +
    '• Thursdays only, and approved before you tell students or families.\n' +
    '• Planned as a team, not by one teacher alone.\n' +
    '• Starts and ends within lunch & recess' + winText + '. It is not an all-day party.\n' +
    '• Include the learning activities for the rest of the day.');
  text(Q.pTitle, 'Leave blank to use “<class> Class Party”.');
  date(Q.pDate, 'Thursdays only.', true);
  time(Q.pStart, 'Within lunch & recess' + winText + '.', true);
  time(Q.pEnd, '', true);
  form.addMultipleChoiceItem().setTitle(Q.pLoc).setChoiceValues(LOCS).showOtherOption(true).setRequired(true);
  text(Q.pTeam, 'Parties are planned as a team. e.g. Ms. Sara, Mr. Omar', true);
  many(Q.pActs, ['Reading Camp', 'Mystery Day', 'Educational games'], true,
    'Reading Camp: camp-themed day built around books · Mystery Day: logic puzzles, hidden pictures, crosswords · Educational games: review current skills', true);
  para(Q.pPlan, 'What will students do before and after lunch & recess?', true);
  para(Q.pFood, 'What will be served, who is bringing it, and how you’ll handle allergies.');
  one(Q.pParents, ['No parent help needed', 'Parent volunteers needed'], false, 'If you need parent volunteers, submit a Volunteer request for them too.');
  many(Q.pConfirm, ['We won’t share this party with students or families until it’s approved.'], true);

  // Page 3 · School event
  const pEvent = form.addPageBreakItem().setTitle('School Event');
  text(Q.eTitle, 'e.g. Classroom Spelling Bee', true);
  date(Q.eDate, '', true);
  time(Q.eStart, '', true);
  time(Q.eEnd, '', true);
  form.addMultipleChoiceItem().setTitle(Q.eLoc).setChoiceValues(LOCS).showOtherOption(true).setRequired(true);
  text(Q.eGrades, 'e.g. 3rd–5th', true);
  para(Q.eDesc, 'Short description and schedule.', true);
  number(Q.eAtt, false);
  text(Q.eBudget, 'e.g. $120 from class fund');
  one(Q.eParents, ['No', 'Yes, on campus', 'Livestream only'], false);
  many(Q.eNeeds, ['Tables & chairs', 'Mic / speakers', 'Projector', 'Livestream', 'Custodial help', 'Front desk support', 'Photos for social media'], false, '', true);

  // Page 4 · Field trip
  const pTrip = form.addPageBreakItem().setTitle('Field Trip');
  text(Q.tDest, '', true);
  text(Q.tAddr, '');
  date(Q.tDate, '', true);
  time(Q.tStart, '', true);
  time(Q.tEnd, '', true);
  text(Q.tGrades, '', true);
  number(Q.tStudents, true);
  text(Q.tChap, 'Names of staff and parent chaperones. Parent chaperones also need a Volunteer request.', true);
  one(Q.tTrans, ['School bus', 'Charter bus', 'Parent drivers', 'Walking'], true);
  text(Q.tCost, 'e.g. $12');
  one(Q.tSlips, ['Ready', 'Need help making them'], false);
  para(Q.tPurpose, 'How does this trip connect to what students are learning?', true);

  // Page 5 · Volunteer / observer
  const pVol = form.addPageBreakItem().setTitle('Volunteer / Observer').setHelpText(
    'Volunteer policy:\n' +
    '• All volunteers and observers need the Principal’s approval in advance.\n' +
    '• High school students and staff members’ students may not be on campus unless approved, and may not have their phones during school hours.\n' +
    '• Once approved, the Front Desk, the Principal and Ms. Omayya are notified automatically.\n' +
    '• Volunteers only get access to what they’re helping with.');
  one(Q.vType, ['Parent volunteer', 'Parent observer', 'High school student', 'Staff member’s student'], true);
  text(Q.vNames, 'e.g. Sr. Amina Khan (Yusuf’s mom)', true);
  text(Q.vContact, 'Optional, for the front desk.');
  date(Q.vDate, '', true);
  date(Q.vEnd, '');
  time(Q.vStart, '', true);
  time(Q.vLeave, '', true);
  para(Q.vTasks, '', true);
  many(Q.vAreas, ['My classroom only', 'Cafeteria', 'Playground', 'Library', 'Gym', 'Front office'], true, '', true);
  many(Q.vConfirm, ['I’ll let them know their access is limited to what they’re helping with (and student volunteers can’t have phones during school hours).'], true);

  // Page 6 · Wrap up
  const pLast = form.addPageBreakItem().setTitle('Almost done');
  para(Q.notes, '');
  text(Q.resubmit, 'Leave blank for a new request. Filled in for you when you use the “Edit & resubmit” button in an email.');

  // Branching: the type question jumps to its page; every type page then skips to the last page.
  typeItem.setChoices([
    typeItem.createChoice('Class Party', pParty),
    typeItem.createChoice('School Event', pEvent),
    typeItem.createChoice('Field Trip', pTrip),
    typeItem.createChoice('Volunteer / Observer', pVol),
  ]);
  pEvent.setGoToPage(pLast);   // after the party page
  pTrip.setGoToPage(pLast);    // after the event page
  pVol.setGoToPage(pLast);     // after the field trip page

  form.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActive().getId());

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'onRequestSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onRequestSubmit').forForm(form).onFormSubmit().create();

  PropertiesService.getScriptProperties().setProperties({ FORM_ID: form.getId(), FORM_URL: form.getPublishedUrl() });
  return form;
}

/* ---------------------------------------------------------------------------
 *  Handle a submission
 * ------------------------------------------------------------------------- */
function onRequestSubmit(e) {
  const resp = e.response;
  const raw = {};
  resp.getItemResponses().forEach(ir => { raw[ir.getItem().getTitle()] = ir.getResponse(); });
  const g = t => (raw[t] == null ? '' : raw[t]);
  const s = t => String(g(t)).trim();
  const answers = pairs => pairs.map(p => ({ key: p[0], label: p[1], value: g(p[1]) }))
    .filter(a => Array.isArray(a.value) ? a.value.length : String(a.value).trim());

  const type = TYPE_CHOICES[s(Q.type)];
  if (!type) return;
  const d = {
    type: type,
    name: s(Q.name),
    grade: s(Q.grade),
    email: String(resp.getRespondentEmail() || '').toLowerCase(),
    raw: raw,
  };

  if (type === 'party') Object.assign(d, {
    title: s(Q.pTitle) || d.grade + ' Class Party',
    date: s(Q.pDate), start: hm_(g(Q.pStart)), end: hm_(g(Q.pEnd)), location: s(Q.pLoc),
    answers: answers([['coPlanners', Q.pTeam], ['activities', Q.pActs], ['learningPlan', Q.pPlan],
      ['food', Q.pFood], ['parentHelp', Q.pParents], ['confirm', Q.pConfirm]]),
  });
  if (type === 'event') Object.assign(d, {
    title: s(Q.eTitle), date: s(Q.eDate), start: hm_(g(Q.eStart)), end: hm_(g(Q.eEnd)), location: s(Q.eLoc),
    answers: answers([['grades', Q.eGrades], ['description', Q.eDesc], ['attendance', Q.eAtt],
      ['budget', Q.eBudget], ['parents', Q.eParents], ['needs', Q.eNeeds]]),
  });
  if (type === 'fieldtrip') Object.assign(d, {
    title: 'Field Trip: ' + s(Q.tDest), date: s(Q.tDate), start: hm_(g(Q.tStart)), end: hm_(g(Q.tEnd)), location: s(Q.tDest),
    answers: answers([['destination', Q.tDest], ['address', Q.tAddr], ['grades', Q.tGrades], ['students', Q.tStudents],
      ['chaperones', Q.tChap], ['transport', Q.tTrans], ['cost', Q.tCost], ['permission', Q.tSlips], ['purpose', Q.tPurpose]]),
  });
  if (type === 'volunteer') Object.assign(d, {
    title: s(Q.vType) + ': ' + s(Q.vNames), date: s(Q.vDate), endDate: s(Q.vEnd),
    start: hm_(g(Q.vStart)), end: hm_(g(Q.vLeave)), location: [].concat(g(Q.vAreas) || []).join(', '),
    answers: answers([['volunteerType', Q.vType], ['volunteerNames', Q.vNames], ['volunteerContact', Q.vContact],
      ['tasks', Q.vTasks], ['areas', Q.vAreas], ['confirm', Q.vConfirm]]),
  });
  if (s(Q.notes)) d.answers.push({ key: 'notes', label: 'Notes', value: s(Q.notes) });

  let r, resubmitted = false;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const prev = s(Q.resubmit) ? getRequest_(s(Q.resubmit).toUpperCase()) : null;
    if (prev && String(prev.email).toLowerCase() === d.email &&
        (prev.status === STATUS.PENDING || prev.status === STATUS.CHANGES)) {
      r = prev;
      resubmitted = true;
    } else {
      r = { id: newId_(), submittedAt: resp.getTimestamp().toISOString(), history: [] };
    }
    Object.assign(r, d, { status: STATUS.PENDING, note: '', decidedBy: '', decidedAt: '' });
    r.history = (r.history || []).concat([{ at: new Date().toISOString(), by: d.name, action: resubmitted ? 'Resubmitted' : 'Submitted' }]);
    saveRequest_(r);
  } finally {
    lock.releaseLock();
  }
  notifyNew_(r, resubmitted);
}

/* ---------------------------------------------------------------------------
 *  Teacher links: "Edit & resubmit" opens the form pre-filled with their answers.
 * ------------------------------------------------------------------------- */
function teacherLink_(r) {
  if (!r || r.status !== STATUS.CHANGES) return '';
  const form = getForm_();
  if (!form) return '';
  const raw = Object.assign({}, r.raw || {});
  raw[Q.resubmit] = r.id;
  let fr = form.createResponse();
  form.getItems().forEach(item => {
    const v = raw[item.getTitle()];
    if (v == null || v === '' || (Array.isArray(v) && !v.length)) return;
    try {
      const T = FormApp.ItemType;
      switch (item.getType()) {
        case T.TEXT: fr = fr.withItemResponse(item.asTextItem().createResponse(String(v))); break;
        case T.PARAGRAPH_TEXT: fr = fr.withItemResponse(item.asParagraphTextItem().createResponse(String(v))); break;
        case T.MULTIPLE_CHOICE: fr = fr.withItemResponse(item.asMultipleChoiceItem().createResponse(String(v))); break;
        case T.CHECKBOX: fr = fr.withItemResponse(item.asCheckboxItem().createResponse([].concat(v))); break;
        case T.DATE: fr = fr.withItemResponse(item.asDateItem().createResponse(ymdToDate_(v))); break;
        case T.TIME: { const p = hm_(v).split(':'); fr = fr.withItemResponse(item.asTimeItem().createResponse(+p[0], +p[1])); break; }
      }
    } catch (err) { /* skip anything that can't be pre-filled */ }
  });
  try { return fr.toPrefilledUrl(); } catch (err) { return form.getPublishedUrl(); }
}
