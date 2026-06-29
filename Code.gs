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
  'Follow Up Completed Date', 'Assigned To', 'Notes',
  'Updated Support Details', 'Follow-Up Outcome', 'Eligibility Review Status',
  'Reviewed By', 'Review Date', 'PI Notes'
];
const FOLLOWUP_REVIEW_FIELDS = [
  'Updated Support Details', 'Follow-Up Outcome', 'Eligibility Review Status',
  'Reviewed By', 'Review Date', 'PI Notes'
];
const MATCH_REVIEW_SHEET = 'Match_Review';
const UNMATCHED_CONSENT_SHEET = 'Unmatched_Consent';
const MASTER_MATCH_FIELDS = [
  'Match Status', 'Match Confidence Score', 'Match Reasons', 'Needs Match Review',
  'Possible Consent Matches', 'Manual Consent ResponseID', 'Manual Match Notes'
];
const MATCH_REVIEW_HEADERS = [
  'EnrollmentID', 'Prescreening ResponseID', 'Child Full Name', 'Parent/Caretaker Name',
  'Parent Email', 'Parent Phone', 'Match Status', 'Match Confidence Score',
  'Match Reasons', 'Possible Consent Matches', 'Manual Consent ResponseID',
  'Manual Match Notes', 'Last Updated'
];
const UNMATCHED_CONSENT_HEADERS = [
  'Consent ResponseID', 'Submitted At', 'Child Full Name', 'Parent Full Name',
  'Parent Email', 'Parent Phone', 'Consent Status', 'Best Prescreening Match',
  'Best Match Score', 'Match Reasons', 'Review Status', 'Notes'
];

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', active.getId());
    return active;
  }
  const storedId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (storedId) return SpreadsheetApp.openById(storedId);
  throw new Error('No spreadsheet is available. Open the Google Sheet once, reload Apps Script, then try the dashboard again.');
}

const DEFAULTS = Object.freeze({
  followUpStatus: 'Not Started',
  consentStatus: 'Pending',
  enrollmentStatus: 'In Progress'
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('G4G Tracker')
    .addItem('Open tracker sidebar', 'showTrackerSidebar')
    .addItem('Open professor dashboard', 'showProfessorDashboard')
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

function showProfessorDashboard() {
  const html = renderDashboardHtml_().setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'G4G Professor Dashboard');
}

/** Returns sidebar-friendly dashboard state after a refresh or on page load. */
function getTrackerSummary() {
  const ss = getSpreadsheet_();
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
  const ss = getSpreadsheet_();
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
  const view = String(e && e.parameter && e.parameter.view || 'dashboard').toLowerCase();
  if (view === 'sidebar') {
    return HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('G4G Enrollment Tracker')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return renderDashboardHtml_();
}

function renderDashboardHtml_() {
  const template = HtmlService.createTemplateFromFile('Dashboard');
  template.initialData = Utilities.base64Encode(JSON.stringify(safeGetAppDashboardData_()));
  return template.evaluate()
    .setTitle('G4G Professor Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function safeGetAppDashboardData_() {
  try {
    return getAppDashboardData();
  } catch (err) {
    return {generatedAt: new Date().toISOString(), error: String(err), metrics: {}, participants: [], followups: [], ready: [], needsReview: [], matchReview: [], unmatchedConsents: [], prescreens: [], consents: [], dashboardRows: [], activity: [], reports: {actionItems: []}, urls: {}, setup: {error: String(err)}};
  }
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
  const ss = getSpreadsheet_();
  const config = getConfig_(ss);
  ensureFollowupReviewColumns_(ss);
  ensureMasterMatchColumns_(ss);
  ensureAuxiliaryMatchSheets_(ss);

  const prescreenRows = readRawRows_(ss.getSheetByName(SHEETS.PRESCREEN_RAW));
  const consentRows = readRawRows_(ss.getSheetByName(SHEETS.CONSENT_RAW));

  const prescreenClean = prescreenRows.map((row, i) => normalizePrescreening_(row, i + 1, config));
  const consentClean = consentRows.map((row, i) => normalizeConsent_(row, i + 1, config));

  writeTable_(ss.getSheetByName(SHEETS.PRESCREEN_CLEAN), prescreenClean, 'ResponseID');
  writeTable_(ss.getSheetByName(SHEETS.CONSENT_CLEAN), consentClean, 'ResponseID');

  const manualFollowupByEnrollment = readManualFollowupState_(ss.getSheetByName(SHEETS.FOLLOWUP));
  const manualMatchByEnrollment = readManualMatchState_(ss.getSheetByName(SHEETS.MASTER));
  const master = buildMasterEnrollment_(prescreenClean, consentClean, manualFollowupByEnrollment, config, manualMatchByEnrollment);
  writeTable_(ss.getSheetByName(SHEETS.MASTER), master, 'EnrollmentID');

  const matchReview = buildMatchReview_(master);
  writeTable_(ss.getSheetByName(MATCH_REVIEW_SHEET), matchReview, 'EnrollmentID');

  const unmatchedConsent = buildUnmatchedConsent_(prescreenClean, consentClean, master);
  writeTable_(ss.getSheetByName(UNMATCHED_CONSENT_SHEET), unmatchedConsent, 'Consent ResponseID');

  const followupQueue = buildFollowupQueue_(master, prescreenClean, manualFollowupByEnrollment, config);
  writeTable_(ss.getSheetByName(SHEETS.FOLLOWUP), followupQueue, 'EnrollmentID');

  const ready = buildReadyForDarts_(master, prescreenClean, consentClean);
  writeTable_(ss.getSheetByName(SHEETS.READY), ready, 'EnrollmentID');
  refreshDashboard();
}


function getProfessorDashboardData() {
  return buildProfessorDashboardData_(false);
}

function refreshProfessorDashboardData() {
  return buildProfessorDashboardData_(true);
}

function getAppDashboardData() {
  return buildAppDashboardData_(false);
}

function refreshAppDashboardData() {
  return buildAppDashboardData_(true);
}

function buildAppDashboardData_(shouldRefresh) {
  if (shouldRefresh) refreshEnrollmentTracker();
  const ss = getSpreadsheet_();
  const professor = buildProfessorDashboardData_(false);
  const dashboardRows = readDashboardMetricRows_(ss);
  const matchReview = readObjectsOrEmpty_(ss, MATCH_REVIEW_SHEET);
  const unmatchedConsents = readObjectsOrEmpty_(ss, UNMATCHED_CONSENT_SHEET);
  const prescreens = readObjectsOrEmpty_(ss, SHEETS.PRESCREEN_CLEAN);
  const consents = readObjectsOrEmpty_(ss, SHEETS.CONSENT_CLEAN);
  const master = readObjectsOrEmpty_(ss, SHEETS.MASTER);
  const followups = readObjectsOrEmpty_(ss, SHEETS.FOLLOWUP);
  const ready = readObjectsOrEmpty_(ss, SHEETS.READY);
  return {
    generatedAt: new Date().toISOString(),
    spreadsheetName: ss.getName(),
    metrics: buildCommandCenterMetrics_(professor.summary, dashboardRows, matchReview, unmatchedConsents),
    dashboardRows,
    participants: professor.participants,
    followups: professor.followups,
    needsReview: professor.needsReview,
    ready: professor.ready,
    matchReview,
    unmatchedConsents,
    prescreens,
    consents,
    master,
    rawCounts: {
      prescreening: readRawRows_(ss.getSheetByName(SHEETS.PRESCREEN_RAW)).length,
      consent: readRawRows_(ss.getSheetByName(SHEETS.CONSENT_RAW)).length
    },
    activity: buildActivityFeed_(prescreens, consents, followups, matchReview, unmatchedConsents),
    reports: buildReportData_(professor.participants, followups, matchReview, unmatchedConsents, ready),
    setup: getSheetSetupStatus_(),
    urls: {
      webApp: getWebAppUrl_(),
      prescreeningWebhook: buildWebhookUrl_('prescreening'),
      consentWebhook: buildWebhookUrl_('consent')
    }
  };
}

function readObjectsOrEmpty_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW) : [];
}

function readDashboardMetricRows_(ss) {
  const sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(2).filter(row => row[0]).map(row => ({
    metric: String(row[0] || ''),
    value: Number(row[1] || 0),
    source: row[2] || '',
    interpretation: row[3] || ''
  }));
}

function buildCommandCenterMetrics_(summary, dashboardRows, matchReview, unmatchedConsents) {
  const metricMap = new Map(dashboardRows.map(row => [row.metric, row.value]));
  return {
    prescreened: metricMap.get('Total Prescreening Submitted') || summary.prescreened || 0,
    consentSubmitted: metricMap.get('Total Consent Submitted') || summary.consentSubmitted || 0,
    masterRecords: metricMap.get('Total Master Enrollment Records') || summary.masterRecords || 0,
    neurodivergentYes: summary.neurodivergentYes || 0,
    followUpNeeded: metricMap.get('Follow Up Needed') || summary.followUpNeeded || 0,
    followUpCompleted: metricMap.get('Follow Up Completed') || summary.followUpCompleted || 0,
    consentCompleted: metricMap.get('Consent Completed') || summary.consentCompleted || 0,
    readyForDarts: metricMap.get('Ready for DARTS') || summary.readyForDarts || 0,
    needsReview: metricMap.get('Needs Review') || summary.needsReview || 0,
    matchReviewNeeded: matchReview.length,
    unmatchedConsent: unmatchedConsents.length
  };
}

function buildActivityFeed_(prescreens, consents, followups, matchReview, unmatchedConsents) {
  const items = [];
  prescreens.slice(-8).forEach(row => items.push({type: 'Prescreening', title: row['Child Full Name'] || 'Prescreening submitted', detail: row['Parent Email'] || row['Parent/Caretaker Name'] || '', when: row['Submitted At'] || ''}));
  consents.slice(-8).forEach(row => items.push({type: 'Consent', title: row['Child Full Name'] || 'Consent submitted', detail: row['Parent Email'] || row['Parent Full Name'] || '', when: row['Submitted At'] || ''}));
  followups.filter(row => row['Follow Up Status'] && row['Follow Up Status'] !== 'Not Started').slice(-8).forEach(row => items.push({type: 'Follow-Up', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Follow Up Status'], when: row['Follow Up Completed Date'] || row['Email Sent Date'] || ''}));
  matchReview.slice(-8).forEach(row => items.push({type: 'Match Review', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Match Status'] || 'Needs Review', when: row['Last Updated'] || ''}));
  unmatchedConsents.slice(-8).forEach(row => items.push({type: 'Unmatched Consent', title: row['Child Full Name'] || row['Consent ResponseID'], detail: row['Best Prescreening Match'] || 'No confident match', when: row['Submitted At'] || ''}));
  return items.sort((a, b) => String(b.when || '').localeCompare(String(a.when || ''))).slice(0, 15);
}

function buildReportData_(participants, followups, matchReview, unmatchedConsents, ready) {
  return {
    followupStatus: groupCounts_(followups, 'Follow Up Status'),
    consentStatus: groupCounts_(participants, 'consentStatus'),
    readiness: groupCounts_(participants, 'readyForDarts'),
    matchStatus: groupCounts_(participants, 'matchStatus'),
    reviewStatus: groupCounts_(participants, 'eligibilityReviewStatus'),
    actionItems: [
      {label: 'Follow-ups not started', value: followups.filter(row => !row['Follow Up Status'] || row['Follow Up Status'] === 'Not Started').length},
      {label: 'Follow-ups awaiting response', value: followups.filter(row => row['Follow Up Status'] === 'Awaiting Response').length},
      {label: 'Match review needed', value: matchReview.length},
      {label: 'Unmatched consent records', value: unmatchedConsents.length},
      {label: 'Ready for DARTS export', value: ready.length}
    ]
  };
}

function groupCounts_(rows, field) {
  return rows.reduce((acc, row) => {
    const key = String(row[field] || 'Missing');
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function approveConsentMatch(update) {
  const ss = getSpreadsheet_();
  ensureMasterMatchColumns_(ss);
  const sheet = ss.getSheetByName(SHEETS.MASTER);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const enrollmentCol = headers.indexOf('EnrollmentID') + 1;
  const prescreenCol = headers.indexOf('Prescreening ResponseID') + 1;
  const manualConsentCol = headers.indexOf('Manual Consent ResponseID') + 1;
  const notesCol = headers.indexOf('Manual Match Notes') + 1;
  if (manualConsentCol < 1 || notesCol < 1) throw new Error('Manual match columns are missing from Master_Enrollment. Refresh the tracker and try again.');
  const enrollmentId = String(update.enrollmentId || '').trim();
  const prescreeningResponseId = String(update.prescreeningResponseId || '').trim();
  const consentResponseId = String(update.consentResponseId || '').trim();
  if (!consentResponseId) throw new Error('Consent ResponseID is required to approve a match.');
  const values = sheet.getDataRange().getValues();
  const targetIndex = values.findIndex((row, index) => index > 0 && ((enrollmentId && String(row[enrollmentCol - 1]) === enrollmentId) || (prescreeningResponseId && String(row[prescreenCol - 1]) === prescreeningResponseId)));
  if (targetIndex < 1) throw new Error('Could not find the selected prescreening/master record to approve this match.');
  const targetRow = targetIndex + 1;
  sheet.getRange(targetRow, manualConsentCol).setValue(consentResponseId);
  sheet.getRange(targetRow, notesCol).setValue(update.notes || `Approved manually on ${new Date().toISOString()}`);
  refreshEnrollmentTracker();
  return buildAppDashboardData_(false);
}


function buildProfessorDashboardData_(shouldRefresh) {
  if (shouldRefresh) refreshEnrollmentTracker();
  const ss = getSpreadsheet_();
  const master = readObjects_(ss.getSheetByName(SHEETS.MASTER), CLEAN_HEADER_ROW, DATA_START_ROW);
  const followups = readObjects_(ss.getSheetByName(SHEETS.FOLLOWUP), CLEAN_HEADER_ROW, DATA_START_ROW);
  const prescreens = readObjects_(ss.getSheetByName(SHEETS.PRESCREEN_CLEAN), CLEAN_HEADER_ROW, DATA_START_ROW);
  const consents = readObjects_(ss.getSheetByName(SHEETS.CONSENT_CLEAN), CLEAN_HEADER_ROW, DATA_START_ROW);
  const ready = readObjects_(ss.getSheetByName(SHEETS.READY), CLEAN_HEADER_ROW, DATA_START_ROW);
  const followupById = new Map(followups.map(row => [String(row['EnrollmentID'] || ''), row]));
  const prescreenByResponse = new Map(prescreens.map(row => [String(row['ResponseID'] || ''), row]));
  const consentByResponse = new Map(consents.map(row => [String(row['ResponseID'] || ''), row]));
  const sourceRecords = master.length ? master : prescreens.map(row => masterLikeFromPrescreen_(row));
  const participants = sourceRecords.map(row => {
    const followup = followupById.get(String(row['EnrollmentID'] || '')) || {};
    const prescreen = prescreenByResponse.get(String(row['Prescreening ResponseID'] || '')) || {};
    const consent = consentByResponse.get(String(row['Consent ResponseID'] || '')) || {};
    return {
      enrollmentId: row['EnrollmentID'] || '',
      childName: row['Child Full Name'] || '',
      parentName: row['Parent/Caretaker Name'] || '',
      parentEmail: row['Parent Email'] || '',
      parentPhone: row['Parent Phone'] || '',
      grade: prescreen['Child Grade'] || consent['Grade'] || '',
      prescreeningStatus: row['Prescreening Status'] || '',
      consentStatus: row['Consent Status'] || '',
      neurodivergentResponse: row['Neurodivergent Response'] || '',
      conditions: followup['Conditions/Diagnoses'] || prescreen['Conditions/Diagnoses'] || '',
      supportDetails: followup['Updated Support Details'] || followup['Support Details'] || prescreen['Diagnostic/Support Details'] || '',
      originalSupportDetails: followup['Support Details'] || prescreen['Diagnostic/Support Details'] || '',
      updatedSupportDetails: followup['Updated Support Details'] || '',
      physicalSupports: followup['Physical Disability Supports'] || prescreen['Physical Disability Supports'] || '',
      followUpNeeded: row['Follow Up Needed'] || '',
      followUpStatus: row['Follow Up Status'] || '',
      emailDrafted: followup['Email Drafted'] || '',
      emailSentDate: followup['Email Sent Date'] || '',
      followUpCompletedDate: followup['Follow Up Completed Date'] || '',
      assignedTo: followup['Assigned To'] || '',
      followUpOutcome: followup['Follow-Up Outcome'] || '',
      eligibilityReviewStatus: followup['Eligibility Review Status'] || deriveReviewStatus_(row, followup, prescreen),
      matchStatus: row['Match Status'] || '',
      matchConfidenceScore: row['Match Confidence Score'] || '',
      matchReasons: row['Match Reasons'] || '',
      needsMatchReview: row['Needs Match Review'] || '',
      possibleConsentMatches: row['Possible Consent Matches'] || '',
      reviewedBy: followup['Reviewed By'] || '',
      reviewDate: followup['Review Date'] || '',
      piNotes: followup['PI Notes'] || '',
      notes: followup['Notes'] || row['Notes'] || '',
      enrollmentStatus: row['Enrollment Status'] || '',
      readyForDarts: row['Ready for DARTS'] || '',
      lastUpdated: row['Last Updated'] || ''
    };
  });
  return {
    generatedAt: new Date().toISOString(),
    summary: buildProfessorSummary_(participants, prescreens, consents, ready, ss),
    participants,
    followups: participants.filter(p => p.followUpNeeded === 'Yes'),
    ready: participants.filter(p => p.readyForDarts === 'Yes'),
    needsReview: participants.filter(p => p.eligibilityReviewStatus === 'Needs Review' || p.readyForDarts === 'Review')
  };
}


function masterLikeFromPrescreen_(prescreen) {
  const followNeeded = prescreen['Follow Up Needed'] || 'No';
  return {
    'EnrollmentID': enrollmentIdFor_(prescreen),
    'Child Full Name': prescreen['Child Full Name'] || '',
    'Parent/Caretaker Name': prescreen['Parent/Caretaker Name'] || '',
    'Parent Email': prescreen['Parent Email'] || '',
    'Parent Phone': prescreen['Parent Phone'] || '',
    'Prescreening Status': prescreen['Response Status'] || 'Completed',
    'Consent Status': 'Pending',
    'Neurodivergent Response': prescreen['Neurodivergent Response'] || '',
    'Follow Up Needed': followNeeded,
    'Follow Up Status': followNeeded === 'Yes' ? (prescreen['Follow Up Status'] || 'Not Started') : 'Not Needed',
    'Enrollment Status': 'In Progress',
    'Ready for DARTS': 'No',
    'Prescreening ResponseID': prescreen['ResponseID'] || '',
    'Consent ResponseID': '',
    'Match Status': 'No Consent Yet',
    'Match Confidence Score': 0,
    'Match Reasons': '',
    'Needs Match Review': 'No',
    'Possible Consent Matches': '',
    'Manual Consent ResponseID': '',
    'Manual Match Notes': '',
    'Last Updated': prescreen['Submitted At'] || ''
  };
}

function buildProfessorSummary_(participants, prescreens, consents, ready, ss) {
  const computed = {
    prescreened: prescreens.length,
    consentSubmitted: consents.length,
    consentCompleted: consents.filter(row => row['Consent Status'] === 'Completed').length,
    masterRecords: participants.length,
    neurodivergentYes: participants.filter(p => p.neurodivergentResponse === 'Yes').length,
    neurodivergentNo: participants.filter(p => p.neurodivergentResponse === 'No').length,
    followUpNeeded: participants.filter(p => p.followUpNeeded === 'Yes').length,
    followUpCompleted: participants.filter(p => p.followUpStatus === 'Completed').length,
    readyForDarts: ready.length,
    needsReview: participants.filter(p => p.eligibilityReviewStatus === 'Needs Review' || p.readyForDarts === 'Review').length
  };
  return Object.assign(computed, readDashboardMetrics_(ss));
}
function readDashboardMetrics_(ss) {
  const sheet = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const metrics = {};
  rows.forEach(row => {
    const name = String(row[0] || '').trim();
    const value = Number(row[1] || 0);
    if (name === 'Total Prescreening Submitted') metrics.prescreened = value;
    if (name === 'Total Consent Submitted') metrics.consentSubmitted = value;
    if (name === 'Total Master Enrollment Records') metrics.masterRecords = value;
    if (name === 'Follow Up Needed') metrics.followUpNeeded = value;
    if (name === 'Follow Up Completed') metrics.followUpCompleted = value;
    if (name === 'Consent Completed') metrics.consentCompleted = value;
    if (name === 'Ready for DARTS') metrics.readyForDarts = value;
    if (name === 'Needs Review') metrics.needsReview = value;
  });
  return metrics;
}

function deriveReviewStatus_(masterRow, followup, prescreen) {
  if (followup['Eligibility Review Status']) return followup['Eligibility Review Status'];
  if (masterRow['Needs Match Review'] === 'Yes') return 'Needs Review';
  if (masterRow['Follow Up Needed'] === 'Yes' && masterRow['Follow Up Status'] !== 'Completed') return 'Needs Review';
  if (masterRow['Consent Status'] === 'Pending') return 'Pending Consent';
  if (masterRow['Ready for DARTS'] === 'Yes') return 'Ready';
  if (prescreen['Diagnostic/Support Details'] && masterRow['Neurodivergent Response'] === 'No') return 'Needs Review';
  return 'Not Reviewed';
}

function updateFollowupReview(updates) {
  const ss = getSpreadsheet_();
  ensureFollowupReviewColumns_(ss);
  const sheet = ss.getSheetByName(SHEETS.FOLLOWUP);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idCol = headers.indexOf('EnrollmentID') + 1;
  const values = sheet.getDataRange().getValues();
  const targetRow = values.findIndex((row, index) => index > 0 && String(row[idCol - 1]) === String(updates.enrollmentId)) + 1;
  if (targetRow < 2) throw new Error(`EnrollmentID not found in Followup_Queue: ${updates.enrollmentId}`);
  const allowed = ['Follow Up Status', 'Email Drafted', 'Email Sent Date', 'Follow Up Completed Date', 'Assigned To', 'Notes', ...FOLLOWUP_REVIEW_FIELDS];
  allowed.forEach(field => {
    if (updates[field] === undefined) return;
    const col = headers.indexOf(field) + 1;
    if (col > 0) sheet.getRange(targetRow, col).setValue(updates[field]);
  });
  refreshEnrollmentTracker();
  return buildAppDashboardData_(false);
}

function ensureFollowupReviewColumns_(ss) {
  const sheet = ss.getSheetByName(SHEETS.FOLLOWUP);
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  FOLLOWUP_REVIEW_FIELDS.forEach(field => {
    if (headers.includes(field)) return;
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(field);
    headers.push(field);
  });
}

function refreshDashboard() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEETS.DASHBOARD);
  const metrics = [
    ['Total Prescreening Submitted', countRows_(ss, SHEETS.PRESCREEN_CLEAN)],
    ['Total Consent Submitted', countRows_(ss, SHEETS.CONSENT_CLEAN)],
    ['Total Master Enrollment Records', countRows_(ss, SHEETS.MASTER)],
    ['Follow Up Needed', countWhere_(ss, SHEETS.MASTER, 'Follow Up Needed', 'Yes')],
    ['Follow Up Sent', countWhere_(ss, SHEETS.FOLLOWUP, 'Follow Up Status', 'Email Sent')],
    ['Follow Up Completed', countWhere_(ss, SHEETS.FOLLOWUP, 'Follow Up Status', 'Completed')],
    ['Consent Completed', countWhere_(ss, SHEETS.MASTER, 'Consent Status', 'Completed')],
    ['Ready for DARTS', countWhere_(ss, SHEETS.MASTER, 'Ready for DARTS', 'Yes')],
    ['Unmatched Consent Records', countRows_(ss, UNMATCHED_CONSENT_SHEET)],
    ['Possible Duplicate Records', countWhere_(ss, MATCH_REVIEW_SHEET, 'Review Status', 'Needs Review')],
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

function buildMasterEnrollment_(prescreens, consents, manualFollowup, config, manualMatchByEnrollment) {
  return prescreens.map(p => {
    const enrollmentId = enrollmentIdFor_(p);
    const manual = manualFollowup.get(enrollmentId) || {};
    const manualMatch = manualMatchByEnrollment.get(enrollmentId) || {};
    const match = selectConsentMatch_(p, consents, manualMatch);
    const consent = match.accepted ? match.consent : null;
    const followStatus = p['Follow Up Needed'] === 'Yes' ? (manual['Follow Up Status'] || p['Follow Up Status'] || config.defaultFollowUpStatus) : 'Not Needed';
    const consentStatus = consent ? consent['Consent Status'] : (match.needsReview ? 'Review' : config.defaultConsentStatus);
    const ready = isReadyWithMatch_(p['Follow Up Needed'], followStatus, consentStatus, match);
    const enrollmentStatus = ready === 'Yes' ? 'Ready for DARTS' : (match.needsReview ? 'Ready for Review' : config.defaultEnrollmentStatus);
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
      'Enrollment Status': enrollmentStatus,
      'Ready for DARTS': ready,
      'Prescreening ResponseID': p['ResponseID'],
      'Consent ResponseID': consent ? consent['ResponseID'] : '',
      'Match Status': match.status,
      'Match Confidence Score': match.score,
      'Match Reasons': match.reasons.join('; '),
      'Needs Match Review': match.needsReview ? 'Yes' : 'No',
      'Possible Consent Matches': match.possibleMatches,
      'Manual Consent ResponseID': manualMatch['Manual Consent ResponseID'] || '',
      'Manual Match Notes': manualMatch['Manual Match Notes'] || '',
      'Last Updated': new Date(),
      'Notes': manual['Notes'] || ''
    };
  });
}


function buildMatchReview_(master) {
  return master.filter(m => m['Needs Match Review'] === 'Yes').map(m => ({
    'EnrollmentID': m['EnrollmentID'],
    'Prescreening ResponseID': m['Prescreening ResponseID'],
    'Child Full Name': m['Child Full Name'],
    'Parent/Caretaker Name': m['Parent/Caretaker Name'],
    'Parent Email': m['Parent Email'],
    'Parent Phone': m['Parent Phone'],
    'Match Status': m['Match Status'],
    'Match Confidence Score': m['Match Confidence Score'],
    'Match Reasons': m['Match Reasons'],
    'Possible Consent Matches': m['Possible Consent Matches'],
    'Manual Consent ResponseID': m['Manual Consent ResponseID'],
    'Manual Match Notes': m['Manual Match Notes'],
    'Review Status': 'Needs Review',
    'Last Updated': m['Last Updated'] || new Date()
  }));
}

function buildUnmatchedConsent_(prescreens, consents, master) {
  const acceptedConsentIds = new Set(master.map(m => String(m['Consent ResponseID'] || '')).filter(Boolean));
  return consents.filter(c => !acceptedConsentIds.has(String(c['ResponseID'] || ''))).map(c => {
    const ranked = prescreens.map(p => ({prescreen: p, match: scoreConsentMatch_(p, c)}))
      .filter(item => item.match.score > 0)
      .sort((a, b) => b.match.score - a.match.score);
    const best = ranked[0];
    return {
      'Consent ResponseID': c['ResponseID'],
      'Submitted At': c['Submitted At'],
      'Child Full Name': c['Child Full Name'],
      'Parent Full Name': c['Parent Full Name'],
      'Parent Email': c['Parent Email'],
      'Parent Phone': c['Parent Phone'],
      'Consent Status': c['Consent Status'],
      'Best Prescreening Match': best ? `${best.prescreen['Child Full Name'] || '(missing child)'} / ${best.prescreen['Parent Email'] || best.prescreen['Parent Phone'] || '(no contact)'} / ${best.prescreen['ResponseID']}` : '',
      'Best Match Score': best ? best.match.score : 0,
      'Match Reasons': best ? best.match.reasons.join('; ') : 'No prescreening record with overlapping child/parent identifiers',
      'Review Status': best && best.match.score >= 40 ? 'Needs Review' : 'Unmatched',
      'Notes': ''
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
      'Notes': manual['Notes'] || '',
      'Updated Support Details': manual['Updated Support Details'] || '',
      'Follow-Up Outcome': manual['Follow-Up Outcome'] || '',
      'Eligibility Review Status': manual['Eligibility Review Status'] || '',
      'Reviewed By': manual['Reviewed By'] || '',
      'Review Date': manual['Review Date'] || '',
      'PI Notes': manual['PI Notes'] || ''
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
  const ss = getSpreadsheet_();
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
function readRawRows_(sheet) {
  return readObjects_(sheet, RAW_HEADER_ROW, RAW_DATA_START_ROW).filter(isMeaningfulRawRow_);
}
function isMeaningfulRawRow_(row) {
  const responseId = String(row['Response ID'] || row['ResponseID'] || '').trim();
  const timestamp = String(row['Timestamp (mm/dd/yyyy)'] || row['Timestamp'] || '').trim();
  const email = String(row['Respondent Email'] || row['Email Address:'] || row['Email Address'] || '').trim();
  const payload = String(row['Custom Variable 5'] || row['External Reference'] || '').trim();
  return [responseId, timestamp, email, payload].some(Boolean);
}

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
  if (!records.length) {
    applyOutputFormatting_(sheet, 0, headers.length);
    return;
  }
  const values = records.map(rec => headers.map(h => rec[h] === undefined ? '' : rec[h]));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  applyOutputFormatting_(sheet, values.length, headers.length);
}

function applyOutputFormatting_(sheet, dataRows, columns) {
  if (![SHEETS.PRESCREEN_CLEAN, SHEETS.FOLLOWUP].includes(sheet.getName())) return;
  const rowsToFormat = Math.max(dataRows, Math.max(0, sheet.getLastRow() - 1), 1);
  const range = sheet.getRange(2, 1, rowsToFormat, columns);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment('middle');
  sheet.setRowHeights(2, rowsToFormat, 24);
}

function ensureMasterMatchColumns_(ss) {
  appendMissingHeaders_(ss.getSheetByName(SHEETS.MASTER), MASTER_MATCH_FIELDS);
}

function ensureAuxiliaryMatchSheets_(ss) {
  ensureSheetWithHeaders_(ss, MATCH_REVIEW_SHEET, MATCH_REVIEW_HEADERS.concat(['Review Status']));
  ensureSheetWithHeaders_(ss, UNMATCHED_CONSENT_SHEET, UNMATCHED_CONSENT_HEADERS);
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  if (sheet.getLastColumn() < headers.length) sheet.insertColumnsAfter(Math.max(sheet.getLastColumn(), 1), headers.length - Math.max(sheet.getLastColumn(), 1));
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0].map(String);
  if (!existing.some(Boolean)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  appendMissingHeaders_(sheet, headers);
  return sheet;
}

function appendMissingHeaders_(sheet, fields) {
  if (!sheet) return;
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  fields.forEach(field => {
    if (headers.includes(field)) return;
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(field);
    headers.push(field);
  });
}

function readManualMatchState_(sheet) {
  const rows = readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW);
  const map = new Map();
  rows.forEach(r => {
    if (!r['EnrollmentID']) return;
    map.set(String(r['EnrollmentID']), {
      'Manual Consent ResponseID': String(r['Manual Consent ResponseID'] || '').trim(),
      'Manual Match Notes': String(r['Manual Match Notes'] || '').trim()
    });
  });
  return map;
}

function selectConsentMatch_(prescreen, consents, manualMatch) {
  const manualConsentId = String(manualMatch['Manual Consent ResponseID'] || '').trim();
  if (manualConsentId) {
    const manualConsent = consents.find(c => String(c['ResponseID'] || '') === manualConsentId);
    if (manualConsent) return {
      consent: manualConsent,
      accepted: true,
      status: 'Manual Match',
      score: 100,
      reasons: ['Manual Consent ResponseID override'],
      needsReview: false,
      possibleMatches: describeConsentMatch_(manualConsent, 100, ['Manual Consent ResponseID override'])
    };
    return {
      consent: null,
      accepted: false,
      status: 'Manual Match Missing',
      score: 0,
      reasons: [`Manual Consent ResponseID not found: ${manualConsentId}`],
      needsReview: true,
      possibleMatches: ''
    };
  }

  const ranked = rankConsentMatches_(prescreen, consents);
  if (!ranked.length) return {
    consent: null,
    accepted: false,
    status: 'No Consent Yet',
    score: 0,
    reasons: ['No submitted consent record shares enough identifiers with this prescreening record'],
    needsReview: false,
    possibleMatches: ''
  };

  const best = ranked[0];
  const second = ranked[1];
  const closeSecond = second && best.score - second.score < 10;
  const strongIdentifier = best.reasons.some(reason => /email exact|phone exact/i.test(reason));
  const accepted = best.score >= 80 && !closeSecond && strongIdentifier;
  const probable = !accepted && best.score >= 60;
  const needsReview = !accepted && best.score >= 40;
  return {
    consent: best.consent,
    accepted,
    status: accepted ? 'Matched' : (probable ? 'Probable Match - Review' : (needsReview ? 'Needs Review' : 'No Consent Yet')),
    score: best.score,
    reasons: closeSecond ? best.reasons.concat(['Another consent record has a similar score; manual review recommended']) : best.reasons,
    needsReview: probable || needsReview || Boolean(closeSecond),
    possibleMatches: ranked.slice(0, 3).map(item => describeConsentMatch_(item.consent, item.score, item.reasons)).join(' | ')
  };
}

function rankConsentMatches_(prescreen, consents) {
  return consents.map(consent => {
    const scored = scoreConsentMatch_(prescreen, consent);
    return {consent, score: scored.score, reasons: scored.reasons};
  }).filter(item => item.score > 0).sort((a, b) => b.score - a.score);
}

function scoreConsentMatch_(prescreen, consent) {
  const reasons = [];
  let score = 0;
  const pChild = nameProfile_(prescreen['Child Full Name']);
  const cChild = nameProfile_(consent['Child Full Name']);
  const pParent = nameProfile_(prescreen['Parent/Caretaker Name']);
  const cParent = nameProfile_(consent['Parent Full Name']);
  const pEmail = normalizeEmail_(prescreen['Parent Email']);
  const cEmail = normalizeEmail_(consent['Parent Email']);
  const pPhone = normalizePhoneDigits_(prescreen['Parent Phone']);
  const cPhone = normalizePhoneDigits_(consent['Parent Phone']);

  if (pChild.full && cChild.full && pChild.full === cChild.full) { score += 40; reasons.push('Child full name exact'); }
  else if (pChild.last && cChild.last && pChild.last === cChild.last && pChild.first && cChild.first && initialsCompatible_(pChild.first, cChild.first)) {
    score += 30; reasons.push('Child first/last name compatible');
  } else {
    const sim = similarity_(pChild.full, cChild.full);
    if (sim >= 0.88) { score += 25; reasons.push(`Child name very similar (${Math.round(sim * 100)}%)`); }
    else if (sim >= 0.75) { score += 15; reasons.push(`Child name somewhat similar (${Math.round(sim * 100)}%)`); }
  }
  if (pEmail && cEmail && pEmail === cEmail) { score += 40; reasons.push('Parent email exact'); }
  if (pPhone && cPhone && pPhone === cPhone) { score += 35; reasons.push('Parent phone exact'); }
  if (pParent.full && cParent.full && pParent.full === cParent.full) { score += 15; reasons.push('Parent/caretaker name exact'); }
  else {
    const parentSim = similarity_(pParent.full, cParent.full);
    if (parentSim >= 0.85) { score += 10; reasons.push(`Parent/caretaker name similar (${Math.round(parentSim * 100)}%)`); }
  }
  if (pChild.last && cChild.last && pChild.last === cChild.last) { score += 5; reasons.push('Child last name exact'); }
  if (gradeCompatible_(prescreen['Child Grade'], consent['Grade'])) { score += 5; reasons.push('Grade compatible'); }
  return {score: Math.min(score, 100), reasons};
}

function describeConsentMatch_(consent, score, reasons) {
  return `${consent['ResponseID'] || '(no response id)'}: ${consent['Child Full Name'] || '(missing child)'} / ${consent['Parent Email'] || consent['Parent Phone'] || '(no contact)'} / score ${score} (${reasons.join(', ')})`;
}

function nameProfile_(value) {
  const full = normalizeNameForMatch_(value);
  const parts = full.split(' ').filter(Boolean);
  return {full, first: parts[0] || '', last: parts.length > 1 ? parts[parts.length - 1] : ''};
}

function normalizeNameForMatch_(value) {
  return String(value || '').toLowerCase().replace(/[^a-z\s]/g, ' ').split(/\s+/).filter(part => part && part.length > 1).join(' ');
}

function initialsCompatible_(a, b) {
  return a === b || (a && b && a[0] === b[0]);
}

function similarity_(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const distance = levenshtein_(a, b);
  return 1 - (distance / Math.max(a.length, b.length));
}

function levenshtein_(a, b) {
  const prev = Array.from({length: b.length + 1}, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const curr = [i];
    for (let j = 1; j <= b.length; j++) {
      curr[j] = Math.min(curr[j - 1] + 1, prev[j] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    for (let j = 0; j < curr.length; j++) prev[j] = curr[j];
  }
  return prev[b.length];
}

function gradeCompatible_(a, b) {
  const left = normalizeToken_(a);
  const right = normalizeToken_(b);
  const leftNumber = (String(a || '').match(/\d+/) || [''])[0];
  const rightNumber = (String(b || '').match(/\d+/) || [''])[0];
  return Boolean(left && right && (left === right || left.replace(/grade$/, '') === right.replace(/grade$/, '') || (leftNumber && leftNumber === rightNumber)));
}

function normalizePhoneDigits_(v) { return String(v || '').replace(/\D/g, '').slice(-10); }
function isReadyWithMatch_(needed, status, consentStatus, match) { return match.accepted && isReady_(needed, status, consentStatus) === 'Yes' ? 'Yes' : (match.needsReview ? 'Review' : 'No'); }

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
function isRawPayloadValue_(value) {
  const text = String(value || '').trim();
  return text.startsWith('{"form"') || text.startsWith('{"responseID"') || text.includes('"responseSet"');
}
function collectOptionValues_(obj, options) {
  const found = [];
  Object.entries(obj).forEach(([k, v]) => {
    if (!v || isRawPayloadValue_(v)) return;
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
  const sh = getSpreadsheet_().getSheetByName(sheetName);
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
function isReady_(needed, status, consentStatus) { return consentStatus === 'Completed' && (needed !== 'Yes' || status === 'Completed') ? 'Yes' : 'No'; }
function followupReason_(p) { return ['Neurodivergent/disability response', p['Conditions/Diagnoses'] && 'Conditions/diagnoses listed', p['Physical Disability Supports'] && 'Physical support needs'].filter(Boolean).join('; '); }
function cleanName_(v) { return String(v || '').trim().replace(/\s+/g, ' ').replace(/\b\w/g, c => c.toUpperCase()); }
function normalizeEmail_(v) { return String(v || '').trim().toLowerCase(); }
function normalizePhone_(v) { const d = String(v || '').replace(/\D/g, ''); return d.length === 10 ? `${d.slice(0,3)}-${d.slice(3,6)}-${d.slice(6)}` : String(v || '').trim(); }
function normalizeToken_(v) { return String(v || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function countRows_(ss, name) { const sheet = ss.getSheetByName(name); return sheet ? readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW).length : 0; }
function countWhere_(ss, name, field, expected) { const sheet = ss.getSheetByName(name); return sheet ? readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW).filter(r => r[field] === expected).length : 0; }
function jsonResponse_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function ensureRequiredSheets_(ss) { Object.values(SHEETS).forEach(n => { if (!ss.getSheetByName(n)) throw new Error(`Missing required sheet: ${n}`); }); }
