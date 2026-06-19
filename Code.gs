/**
 * Gaming4Good Enrollment Tracker — Phase 1 Apps Script automation.
 *
 * Workflow:
 * QuestionPro webhook/manual raw import -> Prescreening_Raw/Consent_Raw audit tabs
 * -> normalized clean tabs -> Master_Enrollment -> Followup_Queue and Ready_For_DARTS.
 *
 * Phase 1 intentionally does not send email. The follow-up queue only identifies
 * who needs contact and lets the study team track drafted/sent/completed status.
 */

const SHEETS = Object.freeze({
  PRESCREEN_RAW: 'Prescreening_Raw',
  CONSENT_RAW: 'Consent_Raw',
  PRESCREEN_CLEAN: 'Prescreening_Clean',
  CONSENT_CLEAN: 'Consent_Clean',
  MASTER: 'Master_Enrollment',
  FOLLOWUP: 'Followup_Queue',
  READY: 'Ready_For_DARTS',
  DASHBOARD: 'Dashboard',
  CONFIG: 'Config'
});

const RAW_HEADER_ROW = 2;
const CLEAN_HEADER_ROW = 1;
const DATA_START_ROW = 2;
const RAW_DATA_START_ROW = 3;
const MANUAL_FOLLOWUP_FIELDS = [
  'Follow Up Status', 'Email Drafted', 'Email Sent Date',
  'Follow Up Completed Date', 'Assigned To', 'Notes'
];
const DEFAULTS = Object.freeze({
  followUpStatus: 'Not Started',
  consentStatus: 'Pending',
  enrollmentStatus: 'In Progress'
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('G4G Tracker')
    .addItem('Open tracker sidebar', 'showTrackerSidebar')
    .addItem('Refresh tracker from raw tabs', 'refreshEnrollmentTracker')
    .addItem('Rebuild dashboard counts', 'refreshDashboard')
    .addSeparator()
    .addItem('Run self-test with sample rows', 'runPhaseOneSelfTest')
    .addToUi();
}


/** Opens a lightweight Sheets sidebar for the Phase 1 tracker controls. */
function showTrackerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('G4G Tracker')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Returns sidebar-friendly dashboard state after a refresh or on page load. */
function getTrackerSummary() {
  const ss = SpreadsheetApp.getActive();
  return {
    spreadsheetName: ss.getName(),
    prescreening: countRows_(ss, SHEETS.PRESCREEN_CLEAN),
    consent: countRows_(ss, SHEETS.CONSENT_CLEAN),
    master: countRows_(ss, SHEETS.MASTER),
    followupNeeded: countWhere_(ss, SHEETS.MASTER, 'Follow Up Needed', 'Yes'),
    followupCompleted: countWhere_(ss, SHEETS.FOLLOWUP, 'Follow Up Status', 'Completed'),
    readyForDarts: countWhere_(ss, SHEETS.MASTER, 'Ready for DARTS', 'Yes'),
    generatedAt: new Date().toISOString(),
    webAppUrl: getWebAppUrl_(),
    prescreeningWebhookUrl: buildWebhookUrl_('prescreening'),
    consentWebhookUrl: buildWebhookUrl_('consent'),
    setup: getSheetSetupStatus_()
  };
}

/** Sidebar wrapper that refreshes the pipeline and returns updated counts. */
function refreshFromSidebar() {
  refreshEnrollmentTracker();
  return getTrackerSummary();
}



function getSheetSetupStatus_() {
  const ss = SpreadsheetApp.getActive();
  const missingSheets = Object.values(SHEETS).filter(name => !ss.getSheetByName(name));
  const requiredRawHeaders = {
    [SHEETS.PRESCREEN_RAW]: ['Response ID', 'Timestamp (mm/dd/yyyy)'],
    [SHEETS.CONSENT_RAW]: ['Response ID', 'Timestamp (mm/dd/yyyy)']
  };
  const missingHeaders = [];
  Object.keys(requiredRawHeaders).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const headers = sheet.getRange(RAW_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    requiredRawHeaders[sheetName].forEach(header => {
      if (!headers.includes(header)) missingHeaders.push(`${sheetName}: ${header}`);
    });
  });
  return {
    ok: missingSheets.length === 0 && missingHeaders.length === 0,
    missingSheets,
    missingHeaders
  };
}

function getWebAppUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (err) {
    return '';
  }
}

function buildWebhookUrl_(formType) {
  const url = getWebAppUrl_();
  return url ? `${url}?form=${encodeURIComponent(formType)}` : '';
}

/** Serves the deployed web app URL. This prevents "Script function not found: doGet". */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('G4G Enrollment Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Receives QuestionPro webhook payloads. Deploy as a web app for webhook use. */
function doPost(e) {
  try {
    const payload = parseWebhookPayload_(e);
    const formType = getRequestedFormType_(e, payload) || detectFormType_(payload);
    appendRawPayload_(formType, payload);
    refreshEnrollmentTracker();
    return jsonResponse_({ok: true, formType, message: `Webhook captured in ${formType} raw tab`});
  } catch (err) {
    console.error(err.stack || err);
    return jsonResponse_({ok: false, error: String(err)});
  }
}

/** Main Phase 1 processing entry point. Safe to rerun; it updates by response/enrollment IDs. */
function refreshEnrollmentTracker() {
  const ss = SpreadsheetApp.getActive();
  const config = getConfig_(ss);

  const prescreenRows = readRawRows_(ss.getSheetByName(SHEETS.PRESCREEN_RAW));
  const consentRows = readRawRows_(ss.getSheetByName(SHEETS.CONSENT_RAW));

  const prescreenClean = prescreenRows.map((row, i) => normalizePrescreening_(row, i + 1, config));
  const consentClean = consentRows.map((row, i) => normalizeConsent_(row, i + 1, config));

  writeTable_(ss.getSheetByName(SHEETS.PRESCREEN_CLEAN), prescreenClean, 'ResponseID');
  writeTable_(ss.getSheetByName(SHEETS.CONSENT_CLEAN), consentClean, 'ResponseID');

  const manualFollowupByEnrollment = readManualFollowupState_(ss.getSheetByName(SHEETS.FOLLOWUP));
  const master = buildMasterEnrollment_(prescreenClean, consentClean, manualFollowupByEnrollment, config);
  writeTable_(ss.getSheetByName(SHEETS.MASTER), master, 'EnrollmentID');

  const followupQueue = buildFollowupQueue_(master, prescreenClean, manualFollowupByEnrollment, config);
  writeTable_(ss.getSheetByName(SHEETS.FOLLOWUP), followupQueue, 'EnrollmentID');

  const ready = buildReadyForDarts_(master, prescreenClean, consentClean);
  writeTable_(ss.getSheetByName(SHEETS.READY), ready, 'EnrollmentID');
  refreshDashboard();
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.DASHBOARD);
  const metrics = [
    ['Total Prescreening Submitted', countRows_(ss, SHEETS.PRESCREEN_CLEAN)],
    ['Total Consent Submitted', countRows_(ss, SHEETS.CONSENT_CLEAN)],
    ['Total Master Enrollment Records', countRows_(ss, SHEETS.MASTER)],
    ['Follow Up Needed', countWhere_(ss, SHEETS.MASTER, 'Follow Up Needed', 'Yes')],
    ['Follow Up Sent', countWhere_(ss, SHEETS.FOLLOWUP, 'Follow Up Status', 'Email Sent')],
    ['Follow Up Completed', countWhere_(ss, SHEETS.FOLLOWUP, 'Follow Up Status', 'Completed')],
    ['Ready for DARTS', countWhere_(ss, SHEETS.MASTER, 'Ready for DARTS', 'Yes')],
    ['Needs Review', countWhere_(ss, SHEETS.MASTER, 'Ready for DARTS', 'Review')]
  ];
  const headerRow = 2;
  const values = sh.getDataRange().getValues();
  const metricCol = 1;
  const valueCol = 2;
  const existing = new Map();
  for (let r = headerRow + 1; r <= values.length; r++) {
    const name = values[r - 1][metricCol - 1];
    if (name) existing.set(String(name), r);
  }
  metrics.forEach(([name, value]) => {
    const row = existing.get(name) || sh.getLastRow() + 1;
    sh.getRange(row, metricCol).setValue(name);
    sh.getRange(row, valueCol).setValue(value);
  });
}

function normalizePrescreening_(raw, index, config) {
  const responseId = getFirst_(raw, ['Response ID', 'ResponseID']) || `P-MANUAL-${index}`;
  const neuroRaw = getByContains_(raw, ['identify as neurodivergent', 'diagnosed disability', 'learning difference']);
  const neuro = decodeQuestionProNeuroResponse_(neuroRaw);
  const conditions = collectOptionValues_(raw, ['Autism', 'ADHD', 'Dyslexia', 'Intellectual disability', 'Developmental disability', 'Epilepsy', 'Traumatic brain injury', 'OCD', 'Down Syndrome', 'Other genetic condition', 'Other (Please specify)']);
  const physicalSupports = getByContains_(raw, ['physically disabled', 'additional supports']);
  const followUpNeeded = needsFollowup_(neuro, conditions, physicalSupports, config) ? 'Yes' : 'No';
  return {
    'PrescreeningID': `PRE-${responseId}`,
    'ResponseID': responseId,
    'Submitted At': getFirst_(raw, ['Timestamp (mm/dd/yyyy)', 'Timestamp']),
    'Response Status': getFirst_(raw, ['Response Status']) || 'Completed',
    'Child Full Name': cleanName_(getByContains_(raw, ['Childs Name', "Child's Name", 'Child Full Name'])),
    'Child Age': getByContains_(raw, ['Childs Age', "Child's Age"]),
    'Child Gender': getByContains_(raw, ['Childs Gender', "Child's Gender"]),
    'Child Grade': getByContains_(raw, ['Childs Grade', "Child's Grade", 'Grade Level']),
    'Race/Ethnicity': getByContains_(raw, ['Race / Ethnicity', 'Race/Ethnicity']),
    'Hispanic/Latino': getByContains_(raw, ['Hispanic or Latino']),
    'Attendance Availability': getByContains_(raw, ['How much of the program', 'able to attend']),
    'Neurodivergent Response': neuro,
    'Conditions/Diagnoses': conditions,
    'Diagnostic/Support Details': getByContains_(raw, ['diagnostic information', 'support services', 'accommodation requests']),
    'Physical Disability Supports': physicalSupports,
    'Parent/Caretaker Name': cleanName_(getByContains_(raw, ['Parent / Caretaker Name', 'Parent/Caretaker Name'])),
    'Parent Phone': normalizePhone_(getByContains_(raw, ['Primary Phone Number', 'phone number'])),
    'Parent Email': normalizeEmail_(getByContains_(raw, ['Email Address', 'Respondent Email'])),
    'Acknowledgement': getByContains_(raw, ['Parent/Guardian Acknowledgement']),
    'Follow Up Needed': followUpNeeded,
    'Follow Up Status': followUpNeeded === 'Yes' ? config.defaultFollowUpStatus : 'Not Needed',
    'Processing Status': 'Processed',
    'Notes': ''
  };
}

function normalizeConsent_(raw, index, config) {
  const responseId = getFirst_(raw, ['Response ID', 'ResponseID']) || `C-MANUAL-${index}`;
  const parentFirst = getByContains_(raw, ['YOUR first name']);
  const parentLast = getByContains_(raw, ['YOUR last name']);
  const childFirst = getByContains_(raw, ["CHILD's first name", 'CHILD first name']);
  const childLast = getByContains_(raw, ["CHILD's last name", 'CHILD last name']);
  const decision = detectConsentDecision_(raw);
  return {
    'ConsentID': `CON-${responseId}`,
    'ResponseID': responseId,
    'Submitted At': getFirst_(raw, ['Timestamp (mm/dd/yyyy)', 'Timestamp']),
    'Response Status': getFirst_(raw, ['Response Status']) || 'Completed',
    'Consent Decision': decision,
    'Parent First Name': cleanName_(parentFirst),
    'Parent Last Name': cleanName_(parentLast),
    'Parent Full Name': cleanName_([parentFirst, parentLast].filter(Boolean).join(' ')),
    'Child First Name': cleanName_(childFirst),
    'Child Last Name': cleanName_(childLast),
    'Child Full Name': cleanName_([childFirst, childLast].filter(Boolean).join(' ')),
    'Relationship to Child': getByContains_(raw, ['relationship to the child']),
    'Parent Phone': normalizePhone_(getByContains_(raw, ['provide your phone number'])),
    'Parent Email': normalizeEmail_(getByContains_(raw, ['future research opportunities', 'email address', 'Respondent Email'])),
    'School': getByContains_(raw, ['school your child attends']),
    'Grade': getByContains_(raw, ['What grade is your child in']),
    'IEP/504': getByContains_(raw, ['IEP', '504 plan']),
    'Neurodivergent Response': decodeQuestionProNeuroResponse_(getByContains_(raw, ['Is your child neurodivergent'])),
    'Relevant Disabilities': collectOptionValues_(raw, ['Autism', 'ADD/ADHD', 'Epilepsy', 'Down Syndrome', 'OCD', 'Generalized Anxiety', 'Other']),
    'Accommodation Needs': getByContains_(raw, ['special needs', 'specific accommodations']),
    'Parent Education': getByContains_(raw, ['highest level of education']),
    'Home Language': getByContains_(raw, ['main language spoken']),
    'Household Income': getByContains_(raw, ['household', 'income']),
    'Consent Status': decision === 'Declined' ? 'Declined' : 'Completed',
    'Processing Status': 'Processed',
    'Notes': ''
  };
}

function buildMasterEnrollment_(prescreens, consents, manualFollowup, config) {
  const consentByKey = new Map(consents.map(c => [matchKey_(c), c]).filter(([k]) => k));
  return prescreens.map(p => {
    const enrollmentId = enrollmentIdFor_(p);
    const consent = consentByKey.get(matchKey_(p)) || findBestConsent_(p, consents);
    const manual = manualFollowup.get(enrollmentId) || {};
    const followStatus = p['Follow Up Needed'] === 'Yes' ? (manual['Follow Up Status'] || p['Follow Up Status'] || config.defaultFollowUpStatus) : 'Not Needed';
    const consentStatus = consent ? consent['Consent Status'] : config.defaultConsentStatus;
    const ready = isReady_(p['Follow Up Needed'], followStatus, consentStatus);
    return {
      'EnrollmentID': enrollmentId,
      'Child Full Name': p['Child Full Name'],
      'Parent/Caretaker Name': p['Parent/Caretaker Name'],
      'Parent Email': p['Parent Email'] || (consent && consent['Parent Email']) || '',
      'Parent Phone': p['Parent Phone'] || (consent && consent['Parent Phone']) || '',
      'Prescreening Status': p['Response Status'] || 'Completed',
      'Consent Status': consentStatus,
      'Neurodivergent Response': p['Neurodivergent Response'],
      'Follow Up Needed': p['Follow Up Needed'],
      'Follow Up Status': followStatus,
      'Enrollment Status': ready === 'Yes' ? 'Ready for DARTS' : config.defaultEnrollmentStatus,
      'Ready for DARTS': ready,
      'Prescreening ResponseID': p['ResponseID'],
      'Consent ResponseID': consent ? consent['ResponseID'] : '',
      'Last Updated': new Date(),
      'Notes': manual['Notes'] || ''
    };
  });
}

function buildFollowupQueue_(master, prescreens, manualFollowup, config) {
  const prescreenById = new Map(prescreens.map(p => [enrollmentIdFor_(p), p]));
  return master.filter(m => m['Follow Up Needed'] === 'Yes').map(m => {
    const p = prescreenById.get(m['EnrollmentID']) || {};
    const manual = manualFollowup.get(m['EnrollmentID']) || {};
    return {
      'EnrollmentID': m['EnrollmentID'],
      'Child Full Name': m['Child Full Name'],
      'Parent/Caretaker Name': m['Parent/Caretaker Name'],
      'Parent Email': m['Parent Email'],
      'Parent Phone': m['Parent Phone'],
      'Neurodivergent Response': m['Neurodivergent Response'],
      'Conditions/Diagnoses': p['Conditions/Diagnoses'],
      'Support Details': p['Diagnostic/Support Details'],
      'Physical Disability Supports': p['Physical Disability Supports'],
      'Follow Up Reason': followupReason_(p),
      'Follow Up Status': manual['Follow Up Status'] || config.defaultFollowUpStatus,
      'Email Drafted': manual['Email Drafted'] || 'No',
      'Email Sent Date': manual['Email Sent Date'] || '',
      'Follow Up Completed Date': manual['Follow Up Completed Date'] || '',
      'Assigned To': manual['Assigned To'] || '',
      'Notes': manual['Notes'] || ''
    };
  });
}

function buildReadyForDarts_(master, prescreens, consents) {
  const prescreenById = new Map(prescreens.map(p => [enrollmentIdFor_(p), p]));
  const consentByResp = new Map(consents.map(c => [c['ResponseID'], c]));
  return master.filter(m => m['Ready for DARTS'] === 'Yes').map(m => {
    const p = prescreenById.get(m['EnrollmentID']) || {};
    const c = consentByResp.get(m['Consent ResponseID']) || {};
    return {
      'EnrollmentID': m['EnrollmentID'],
      'Child Full Name': m['Child Full Name'],
      'Parent Name': m['Parent/Caretaker Name'],
      'Parent Email': m['Parent Email'],
      'Parent Phone': m['Parent Phone'],
      'Grade': p['Child Grade'] || c['Grade'],
      'Neurodivergent Response': m['Neurodivergent Response'],
      'Conditions/Diagnoses': p['Conditions/Diagnoses'],
      'Consent Status': m['Consent Status'],
      'Follow Up Status': m['Follow Up Status'],
      'Enrollment Status': m['Enrollment Status'],
      'Ready for DARTS': m['Ready for DARTS']
    };
  });
}

function runPhaseOneSelfTest() {
  const ss = SpreadsheetApp.getActive();
  ensureRequiredSheets_(ss);
  appendRawObjectForTest_(SHEETS.PRESCREEN_RAW, {
    'Response ID': `TEST-P-${Date.now()}`,
    'Response Status': 'Completed',
    'Timestamp (mm/dd/yyyy)': new Date(),
    'Childs Name (First Last)': 'Sample Child',
    'Does your child identify as neurodivergent or have a diagnosed disability developmental condition or learning difference?': 'Yes',
    'Autism': 'Autism',
    'Please provide as much detail as youre willing to share regarding diagnostic information current support services and needs accommodation requests and other relevant information.': 'Needs sensory supports.',
    'Parent / Caretaker Name (First Last):': 'Sample Parent',
    'Primary Phone Number. Note that your phone number should be 10-digits and in this format. XXX-XXX-XXXX': '555-555-1212',
    'Email Address:': 'sample.parent@example.org'
  });
  refreshEnrollmentTracker();
}

// ---------- Helpers ----------
function getConfig_(ss) {
  const rows = ss.getSheetByName(SHEETS.CONFIG).getDataRange().getValues();
  const map = new Map(rows.slice(1).map(r => [String(r[0] || '').trim(), r[1]]));
  return {
    triggerResponse: String(map.get('Follow-up Trigger Response') || 'Yes'),
    defaultFollowUpStatus: String(map.get('Default Follow-up Status') || DEFAULTS.followUpStatus),
    defaultConsentStatus: String(map.get('Default Consent Status') || DEFAULTS.consentStatus),
    defaultEnrollmentStatus: String(map.get('Default Enrollment Status') || DEFAULTS.enrollmentStatus)
  };
}
function readRawRows_(sheet) { return readObjects_(sheet, RAW_HEADER_ROW, RAW_DATA_START_ROW).filter(r => Object.values(r).some(Boolean)); }
function readObjects_(sheet, headerRow, dataStartRow) {
  const values = sheet.getDataRange().getValues();
  const subheaders = headerRow + 1 <= values.length ? values[headerRow] : [];
  const headers = (values[headerRow - 1] || []).map((h, i) => String(h || subheaders[i] || '').trim());
  return values.slice(dataStartRow - 1).filter(r => r.some(Boolean)).map(row => {
    const obj = {};
    headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
    return obj;
  });
}
function writeTable_(sheet, records, keyField) {
  const headers = sheet.getRange(CLEAN_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  if (!records.length) return;
  const values = records.map(rec => headers.map(h => rec[h] === undefined ? '' : rec[h]));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}
function readManualFollowupState_(sheet) {
  const rows = readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW);
  const map = new Map();
  rows.forEach(r => {
    if (!r['EnrollmentID']) return;
    const state = {};
    MANUAL_FOLLOWUP_FIELDS.forEach(f => state[f] = r[f] || '');
    map.set(String(r['EnrollmentID']), state);
  });
  return map;
}
function getFirst_(obj, keys) { for (const k of keys) if (obj[k]) return obj[k]; return ''; }
function getByContains_(obj, needles) {
  const lower = needles.map(n => String(n).toLowerCase());
  for (const [k, v] of Object.entries(obj)) if (v && lower.every(n => String(k).toLowerCase().includes(n))) return v;
  for (const [k, v] of Object.entries(obj)) if (v && lower.some(n => String(k).toLowerCase().includes(n))) return v;
  return '';
}
function collectOptionValues_(obj, options) {
  const found = [];
  Object.entries(obj).forEach(([k, v]) => {
    if (!v) return;
    const hit = options.find(o => String(k).toLowerCase().includes(String(o).toLowerCase()) || String(v).toLowerCase().includes(String(o).toLowerCase()));
    if (hit) found.push(String(v) === '1' || String(v).toLowerCase() === 'yes' ? hit : String(v));
  });
  return [...new Set(found)].join('; ');
}

/**
 * QuestionPro stores the prescreening neurodivergence question as numeric choices:
 * 1 = No (child is not neurodivergent), 2 = Yes (child is neurodivergent).
 * Preserve readable Yes/No values in clean/master tabs so follow-up decisions are correct.
 */
function decodeQuestionProNeuroResponse_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized === '1') return 'No';
  if (normalized === '2') return 'Yes';
  if (/^no\b/.test(normalized)) return 'No';
  if (/^yes\b/.test(normalized)) return 'Yes';
  return value || '';
}

function needsFollowup_(neuro, conditions, physicalSupports, config) {
  return [neuro, conditions, physicalSupports].some(v => /\byes\b|autism|adhd|dyslexia|disability|iep|504|support|accommodation|diagnos/i.test(String(v || '')));
}
function detectConsentDecision_(raw) {
  const text = Object.values(raw).join(' ').toLowerCase();
  if (/do not consent|decline|no,? i do not/.test(text)) return 'Declined';
  return 'Consented';
}
function detectFormType_(payload) {
  const text = JSON.stringify(payload).toLowerCase();
  if (text.includes('consent') || text.includes("child's first name") || text.includes('parental/guardian consent')) return 'Consent';
  return 'Prescreening';
}
function parseWebhookPayload_(e) {
  if (!e) return {};
  const payload = {};
  Object.assign(payload, e.parameter || {});
  const contents = e.postData && e.postData.contents ? String(e.postData.contents).trim() : '';
  if (!contents) return payload;
  try {
    const parsed = JSON.parse(contents);
    return Object.assign(payload, parsed);
  } catch (err) {
    // QuestionPro may send form-encoded bodies instead of raw JSON.
    if (contents.includes('=')) Object.assign(payload, parseFormEncoded_(contents));
    else payload.rawBody = contents;
    return payload;
  }
}
function getRequestedFormType_(e, payload) {
  const requested = String((e && e.parameter && (e.parameter.form || e.parameter.formType || e.parameter.surveyType)) || payload.form || payload.formType || payload.surveyType || '').toLowerCase();
  if (/consent/.test(requested)) return 'Consent';
  if (/pre|screen|interest/.test(requested)) return 'Prescreening';
  return '';
}
function parseFormEncoded_(contents) {
  return contents.split('&').reduce((obj, pair) => {
    const [rawKey, rawValue = ''] = pair.split('=');
    const key = decodeURIComponent(rawKey.replace(/\+/g, ' '));
    obj[key] = decodeURIComponent(rawValue.replace(/\+/g, ' '));
    return obj;
  }, {});
}
function appendRawPayload_(formType, payload) {
  const flattened = flattenPayload_(payload);
  Object.assign(flattened, extractQuestionProResponseSet_(payload));
  removeRoutingOnlyFields_(flattened);
  flattened._rawPayload = JSON.stringify(payload);
  appendRawObjectForTest_(formType === 'Consent' ? SHEETS.CONSENT_RAW : SHEETS.PRESCREEN_RAW, flattened);
}

function removeRoutingOnlyFields_(obj) {
  ['form', 'formType', 'surveyType'].forEach(key => delete obj[key]);
}
function extractQuestionProResponseSet_(payload) {
  const out = {};
  const responseSet = Array.isArray(payload.responseSet) ? payload.responseSet : [];
  responseSet.forEach(item => {
    const key = item.questionText || item.questionDescription || item.questionCode || item.questionID;
    const value = extractQuestionProAnswer_(item);
    if (key && value !== '') out[key] = value;
    if (item.questionCode && value !== '') out[item.questionCode] = value;
  });
  return out;
}
function extractQuestionProAnswer_(item) {
  const directFields = ['answerText', 'answer', 'responseText', 'response', 'value', 'selectedAnswer', 'displayText'];
  for (const field of directFields) {
    const value = normalizeCellValue_(item[field]);
    if (value !== '') return value;
  }
  const arrayFields = ['answers', 'answerValues', 'values', 'selectedAnswers'];
  for (const field of arrayFields) {
    if (!Array.isArray(item[field])) continue;
    const values = item[field].map(normalizeCellValue_).filter(Boolean);
    if (values.length) return values.join('; ');
  }
  return '';
}
function normalizeCellValue_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value;
  if (Array.isArray(value)) return value.map(normalizeCellValue_).filter(Boolean).join('; ');
  if (typeof value !== 'object') return String(value).trim();
  const answerFields = ['answerText', 'text', 'value', 'label', 'displayText', 'optionText', 'option', 'code', 'name', 'email', 'phone'];
  for (const field of answerFields) {
    if (value[field] === value) continue;
    const normalized = normalizeCellValue_(value[field]);
    if (normalized !== '') return normalized;
  }
  const primitiveValues = Object.values(value).map(normalizeCellValue_).filter(Boolean);
  return primitiveValues.length === 1 ? primitiveValues[0] : '';
}
function appendRawObjectForTest_(sheetName, obj) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const headers = sh.getRange(RAW_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  const normalizedObj = normalizeObjectKeys_(obj);
  const row = headers.map(h => valueForHeader_(h, obj, normalizedObj));
  ensureRawFallbackValues_(headers, row, obj);
  sh.appendRow(row);
}
function normalizeObjectKeys_(obj) {
  const out = {};
  Object.entries(obj || {}).forEach(([key, value]) => {
    if (isQuestionProStructuralKey_(key)) return;
    out[normalizeHeaderKey_(key)] = value;
  });
  return out;
}
function isQuestionProStructuralKey_(key) {
  return /^responseSet\.\d+\.(questionText|questionDescription|questionCode|questionID)$/i.test(String(key || ''));
}
function valueForHeader_(header, obj, normalizedObj) {
  if (!header) return '';
  if (obj[header] !== undefined) return normalizeCellValue_(obj[header]);
  const normalizedHeader = normalizeHeaderKey_(header);
  if (normalizedObj[normalizedHeader] !== undefined) return normalizeCellValue_(normalizedObj[normalizedHeader]);
  if (normalizedHeader.length < 8) return '';
  const partialKey = Object.keys(normalizedObj).find(key => key.length >= 8 && (key.includes(normalizedHeader) || normalizedHeader.includes(key)));
  return partialKey ? normalizeCellValue_(normalizedObj[partialKey]) : '';
}
function ensureRawFallbackValues_(headers, row, obj) {
  setRawFallback_(headers, row, ['Response ID', 'ResponseID', 'responseId'], obj['Response ID'] || obj.responseId || obj.id || `WEBHOOK-${Date.now()}`);
  setRawFallback_(headers, row, ['Response Status'], obj['Response Status'] || obj.status || 'Completed');
  setRawFallback_(headers, row, ['Timestamp (mm/dd/yyyy)', 'Timestamp'], obj['Timestamp (mm/dd/yyyy)'] || obj.timestamp || new Date());
  setRawFallback_(headers, row, ['Custom Variable 5', 'External Reference'], obj._rawPayload || JSON.stringify(obj || {}));
}
function setRawFallback_(headers, row, possibleHeaders, value) {
  const index = headers.findIndex(h => possibleHeaders.some(name => normalizeHeaderKey_(h) === normalizeHeaderKey_(name)));
  if (index >= 0 && !row[index]) row[index] = value;
}
function flattenPayload_(payload, prefix = '', out = {}) {
  Object.entries(payload || {}).forEach(([k, v]) => {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === 'object' && !Array.isArray(v)) flattenPayload_(v, key, out);
    else out[key] = Array.isArray(v) ? v.join('; ') : v;
  });
  return out;
}
function normalizeHeaderKey_(value) { return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function enrollmentIdFor_(p) { return `ENR-${matchKey_(p) || normalizeToken_(p['ResponseID'])}`.toUpperCase(); }
function matchKey_(r) { return [normalizeToken_(r['Child Full Name']), normalizeEmail_(r['Parent Email']) || normalizePhone_(r['Parent Phone'])].filter(Boolean).join('|'); }
function findBestConsent_(p, consents) { return consents.find(c => normalizeToken_(c['Child Full Name']) === normalizeToken_(p['Child Full Name'])) || null; }
function isReady_(needed, status, consentStatus) { return consentStatus === 'Completed' && (needed !== 'Yes' || status === 'Completed') ? 'Yes' : 'No'; }
function followupReason_(p) { return ['Neurodivergent/disability response', p['Conditions/Diagnoses'] && 'Conditions/diagnoses listed', p['Physical Disability Supports'] && 'Physical support needs'].filter(Boolean).join('; '); }
function cleanName_(v) { return String(v || '').trim().replace(/\s+/g, ' ').replace(/\b\w/g, c => c.toUpperCase()); }
function normalizeEmail_(v) { return String(v || '').trim().toLowerCase(); }
function normalizePhone_(v) { const d = String(v || '').replace(/\D/g, ''); return d.length === 10 ? `${d.slice(0,3)}-${d.slice(3,6)}-${d.slice(6)}` : String(v || '').trim(); }
function normalizeToken_(v) { return String(v || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function countRows_(ss, name) { return Math.max(0, ss.getSheetByName(name).getLastRow() - 1); }
function countWhere_(ss, name, field, expected) { return readObjects_(ss.getSheetByName(name), CLEAN_HEADER_ROW, DATA_START_ROW).filter(r => r[field] === expected).length; }
function jsonResponse_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function ensureRequiredSheets_(ss) { Object.values(SHEETS).forEach(n => { if (!ss.getSheetByName(n)) throw new Error(`Missing required sheet: ${n}`); }); }
