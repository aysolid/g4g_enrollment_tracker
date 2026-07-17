/**
 * Gaming4Good Enrollment Tracker — Phase 1 Apps Script automation.
 *
 * Workflow:
 * QuestionPro webhook/manual raw import -> Prescreening_Raw/Consent_Raw audit tabs
 * -> normalized clean tabs -> Screening_Review -> Eligible/Ineligible participant lists
 * -> Master_Enrollment, Match_Review, Followup_Queue, and Ready_For_DARTS.
 *
 * Human-in-the-loop eligibility is intentional. The neurodivergent/support-needs
 * answer is a system suggestion only; study staff make the final eligibility decision.
 * Email automation has been removed from the active workflow.
 */

const SHEETS = Object.freeze({
  PRESCREEN_RAW: 'Prescreening_Raw',
  CONSENT_RAW: 'Consent_Raw',
  PRESCREEN_CLEAN: 'Prescreening_Clean',
  CONSENT_CLEAN: 'Consent_Clean',
  SCREENING_REVIEW: 'Screening_Review',
  ELIGIBLE: 'Eligible_Participants',
  INELIGIBLE: 'Ineligible_Participants',
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
  'Manual Contact Needed', 'Manual Contact Status', 'Assigned To', 'Notes',
  'Updated Support Details', 'Parent Response Summary', 'Eligibility Review Status',
  'Human Eligibility Decision', 'Decision Reason', 'Reviewed By', 'Review Date', 'PI Notes'
];
const FOLLOWUP_REVIEW_FIELDS = [
  'Updated Support Details', 'Parent Response Summary', 'Eligibility Review Status',
  'Human Eligibility Decision', 'Decision Reason', 'Reviewed By', 'Review Date', 'PI Notes'
];
const SCREENING_MANUAL_FIELDS = [
  'Child Full Name', 'Parent/Caretaker Name', 'Parent Email', 'Parent Phone',
  'Cohort ID', 'Cohort Name', 'Site', 'Program Term',
  'Human Eligibility Decision', 'Decision Reason', 'Manual Contact Needed', 'Manual Contact Status',
  'Assigned To', 'Updated Support Details', 'Parent Response Summary', 'Eligibility Review Status',
  'Reviewed By', 'Review Date', 'PI Notes', 'Notes'
];
const SCREENING_REVIEW_HEADERS = [
  'EnrollmentID', 'PrescreeningID', 'ResponseID', 'Submitted At', 'Child Full Name',
  'Parent/Caretaker Name', 'Parent Email', 'Parent Phone', 'Cohort ID', 'Cohort Name',
  'Site', 'Program Term', 'Neurodivergent Response', 'Conditions/Diagnoses',
  'Diagnostic/Support Details', 'Physical Disability Supports', 'System Eligibility Suggestion',
  'System Eligibility Reason', 'Human Eligibility Decision', 'Decision Reason',
  'Manual Contact Needed', 'Manual Contact Status', 'Assigned To', 'Updated Support Details',
  'Parent Response Summary', 'Eligibility Review Status', 'Reviewed By', 'Review Date',
  'PI Notes', 'Notes', 'Last Updated'
];
const ELIGIBLE_PARTICIPANTS_HEADERS = SCREENING_REVIEW_HEADERS.concat(['Eligibility Approved At']);
const INELIGIBLE_PARTICIPANTS_HEADERS = SCREENING_REVIEW_HEADERS.concat(['Eligibility Closed At']);
const MATCH_REVIEW_SHEET = 'Match_Review';
const UNMATCHED_CONSENT_SHEET = 'Unmatched_Consent';
const COHORTS_SHEET = 'Cohorts';
const COHORT_FIELDS = ['Cohort ID', 'Cohort Name', 'Site', 'Program Term'];
const COHORT_HEADERS = ['Cohort ID', 'Cohort Name', 'Site', 'Program Term', 'Status', 'Prescreening Webhook URL', 'Notes', 'Last Seen'];
const MASTER_MATCH_FIELDS = [
  'Match Status', 'Match Confidence Score', 'Match Reasons', 'Needs Match Review',
  'Possible Consent Matches', 'Manual Consent ResponseID', 'Manual Match Notes'
];
const MATCH_REVIEW_HEADERS = [
  'EnrollmentID', 'Prescreening ResponseID', 'Child Full Name', 'Parent/Caretaker Name',
  'Parent Email', 'Parent Phone', 'Cohort ID', 'Cohort Name', 'Site', 'Program Term',
  'Match Status', 'Match Confidence Score',
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
    .addItem('Install raw-tab automation', 'installRawTabAutomation')
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
    screeningReview: countRows_(ss, SHEETS.SCREENING_REVIEW),
    eligible: countRows_(ss, SHEETS.ELIGIBLE),
    ineligible: countRows_(ss, SHEETS.INELIGIBLE),
    master: countRows_(ss, SHEETS.MASTER),
    followupNeeded: countWhere_(ss, SHEETS.SCREENING_REVIEW, 'Manual Contact Needed', 'Yes'),
    followupCompleted: countWhere_(ss, SHEETS.SCREENING_REVIEW, 'Manual Contact Status', 'Completed'),
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

function buildWebhookUrl_(formType, cohort) {
  const url = getWebAppUrl_();
  if (!url) return '';
  const params = [`form=${encodeURIComponent(formType)}`];
  if (cohort && cohort.cohortId) params.push(`cohort_id=${encodeURIComponent(cohort.cohortId)}`);
  if (cohort && cohort.cohortName) params.push(`cohort_name=${encodeURIComponent(cohort.cohortName)}`);
  if (cohort && cohort.site) params.push(`site=${encodeURIComponent(cohort.site)}`);
  if (cohort && cohort.programTerm) params.push(`program_term=${encodeURIComponent(cohort.programTerm)}`);
  return `${url}?${params.join('&')}`;
}

/** Serves the deployed web app URL. This prevents "Script function not found: doGet". */
function doGet(e) {
  const action = String(e && e.parameter && e.parameter.action || '').toLowerCase();
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
    return {generatedAt: new Date().toISOString(), error: String(err), metrics: {}, participants: [], screening: [], eligibleParticipants: [], ineligibleParticipants: [], followups: [], ready: [], needsReview: [], matchReview: [], unmatchedConsents: [], prescreens: [], consents: [], dashboardRows: [], activity: [], reports: {actionItems: []}, urls: {}, setup: {error: String(err)}};
  }
}

/** Receives QuestionPro webhook payloads. Deploy as a web app for webhook use. */
function doPost(e) {
  try {
    const payload = parseWebhookPayload_(e);
    const action = String((e && e.parameter && e.parameter.action) || payload.action || '').toLowerCase();
    const formType = getRequestedFormType_(e, payload) || detectFormType_(payload);
    const cohort = getRequestedCohort_(e, payload);
    appendRawPayload_(formType, payload, cohort);
    upsertCohortFromWebhook_(cohort);
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
  ensureCohortInfrastructure_(ss);
  ensureScreeningInfrastructure_(ss);

  const prescreenRawSheet = ss.getSheetByName(SHEETS.PRESCREEN_RAW);
  const consentRawSheet = ss.getSheetByName(SHEETS.CONSENT_RAW);
  const prescreenRows = readRawRows_(prescreenRawSheet);
  const consentRows = readRawRows_(consentRawSheet);
  applyRawSheetFormatting_(prescreenRawSheet);
  applyRawSheetFormatting_(consentRawSheet);

  const prescreenClean = prescreenRows.map((row, i) => normalizePrescreening_(row, i + 1, config));
  const consentClean = consentRows.map((row, i) => normalizeConsent_(row, i + 1, config));
  ensureCoreOutputHeaders_(ss);

  writeTable_(ss.getSheetByName(SHEETS.PRESCREEN_CLEAN), prescreenClean, 'ResponseID');
  writeTable_(ss.getSheetByName(SHEETS.CONSENT_CLEAN), consentClean, 'ResponseID');

  const manualScreeningByEnrollment = readManualScreeningState_(ss.getSheetByName(SHEETS.SCREENING_REVIEW));
  const legacyFollowupByEnrollment = readManualFollowupState_(ss.getSheetByName(SHEETS.FOLLOWUP));
  const screeningReview = buildScreeningReview_(prescreenClean, manualScreeningByEnrollment, legacyFollowupByEnrollment);
  writeTable_(ss.getSheetByName(SHEETS.SCREENING_REVIEW), screeningReview, 'EnrollmentID');

  const eligible = buildEligibleParticipants_(screeningReview);
  const ineligible = buildIneligibleParticipants_(screeningReview);
  writeTable_(ss.getSheetByName(SHEETS.ELIGIBLE), eligible, 'EnrollmentID');
  writeTable_(ss.getSheetByName(SHEETS.INELIGIBLE), ineligible, 'EnrollmentID');

  const manualMatchByEnrollment = readManualMatchState_(ss.getSheetByName(SHEETS.MASTER));
  const master = buildMasterEnrollment_(eligible, consentClean, config, manualMatchByEnrollment);
  writeTable_(ss.getSheetByName(SHEETS.MASTER), master, 'EnrollmentID');

  const matchReview = buildMatchReview_(master);
  writeTable_(ss.getSheetByName(MATCH_REVIEW_SHEET), matchReview, 'EnrollmentID');

  const unmatchedConsent = buildUnmatchedConsent_(eligible, consentClean, master);
  writeTable_(ss.getSheetByName(UNMATCHED_CONSENT_SHEET), unmatchedConsent, 'Consent ResponseID');

  const followupQueue = buildManualContactQueue_(screeningReview);
  writeTable_(ss.getSheetByName(SHEETS.FOLLOWUP), followupQueue, 'EnrollmentID');

  const ready = buildReadyForDarts_(master, eligible, consentClean);
  writeTable_(ss.getSheetByName(SHEETS.READY), ready, 'EnrollmentID');
  refreshDashboard();
  rememberRawTabSignatures_(ss);
}


function getProfessorDashboardData() {
  return jsonSafe_(buildProfessorDashboardData_(false));
}

function refreshProfessorDashboardData() {
  return jsonSafe_(buildProfessorDashboardData_(true));
}

function getAppDashboardData() {
  return buildAppDashboardData_(false);
}

function refreshAppDashboardData() {
  // Web-app reloads must be read-only. The Google Sheet remains the source of truth.
  return buildAppDashboardData_(false);
}

function buildAppDashboardData_(shouldRefresh) {
  // The dashboard is read-only against the current workbook state.
  // Pipeline rebuilds happen from webhooks, the sidebar/menu, or installed raw-tab triggers.
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
  const screening = readObjectsOrEmpty_(ss, SHEETS.SCREENING_REVIEW);
  const eligibleParticipants = readObjectsOrEmpty_(ss, SHEETS.ELIGIBLE);
  const ineligibleParticipants = readObjectsOrEmpty_(ss, SHEETS.INELIGIBLE);
  const payload = {
    generatedAt: new Date().toISOString(),
    spreadsheetName: ss.getName(),
    metrics: buildCommandCenterMetrics_(professor.summary, dashboardRows, matchReview, unmatchedConsents, screening, eligibleParticipants, ineligibleParticipants),
    dashboardRows,
    participants: professor.participants,
    screening,
    eligibleParticipants,
    ineligibleParticipants,
    followups: professor.followups,
    needsReview: professor.needsReview,
    ready: professor.ready,
    matchReview,
    unmatchedConsents,
    prescreens,
    consents,
    master,
    cohorts: buildCohortSummaries_(ss, professor.participants, prescreens, consents),
    rawCounts: {
      prescreening: readRawRows_(ss.getSheetByName(SHEETS.PRESCREEN_RAW)).length,
      consent: readRawRows_(ss.getSheetByName(SHEETS.CONSENT_RAW)).length
    },
    activity: buildActivityFeed_(prescreens, consents, followups, matchReview, unmatchedConsents, screening),
    reports: buildReportData_(professor.participants, followups, matchReview, unmatchedConsents, ready, screening, eligibleParticipants, ineligibleParticipants),
    setup: getSheetSetupStatus_(),
    urls: {
      webApp: getWebAppUrl_(),
      prescreeningWebhook: buildWebhookUrl_('prescreening'),
      consentWebhook: buildWebhookUrl_('consent')
    }
  };
  return jsonSafe_(payload);
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

function buildCommandCenterMetrics_(summary, dashboardRows, matchReview, unmatchedConsents, screening, eligible, ineligible) {
  const metricMap = new Map(dashboardRows.map(row => [row.metric, row.value]));
  return {
    prescreened: metricMap.get('Total Prescreening Submitted') || summary.prescreened || 0,
    consentSubmitted: metricMap.get('Total Consent Submitted') || summary.consentSubmitted || 0,
    masterRecords: metricMap.get('Total Master Enrollment Records') || summary.masterRecords || 0,
    screeningPending: (screening || []).filter(row => !row['Human Eligibility Decision'] || row['Human Eligibility Decision'] === 'Pending Review').length,
    eligibleApproved: (eligible || []).length,
    ineligible: (ineligible || []).length,
    neurodivergentYes: summary.neurodivergentYes || 0,
    neurodivergentNo: summary.neurodivergentNo || 0,
    manualContactNeeded: metricMap.get('Manual Contact Needed') || summary.followUpNeeded || 0,
    manualContactCompleted: metricMap.get('Manual Contact Completed') || summary.followUpCompleted || 0,
    consentCompleted: metricMap.get('Consent Completed') || summary.consentCompleted || 0,
    readyForDarts: metricMap.get('Ready for DARTS') || summary.readyForDarts || 0,
    needsReview: metricMap.get('Needs Review') || summary.needsReview || 0,
    matchReviewNeeded: matchReview.length,
    unmatchedConsent: unmatchedConsents.length
  };
}

function buildActivityFeed_(prescreens, consents, followups, matchReview, unmatchedConsents, screening) {
  const items = [];
  prescreens.slice(-8).forEach(row => items.push({type: 'Prescreening', title: row['Child Full Name'] || 'Prescreening submitted', detail: row['Parent Email'] || row['Parent/Caretaker Name'] || '', when: row['Submitted At'] || ''}));
  consents.slice(-8).forEach(row => items.push({type: 'Consent', title: row['Child Full Name'] || 'Consent submitted', detail: row['Parent Email'] || row['Parent Full Name'] || '', when: row['Submitted At'] || ''}));
  (screening || []).filter(row => row['Human Eligibility Decision'] && row['Human Eligibility Decision'] !== 'Pending Review').slice(-8).forEach(row => items.push({type: 'Eligibility', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Human Eligibility Decision'], when: row['Review Date'] || row['Last Updated'] || ''}));
  followups.filter(row => row['Follow Up Status'] && row['Follow Up Status'] !== 'Not Started').slice(-8).forEach(row => items.push({type: 'Manual Contact', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Follow Up Status'], when: row['Review Date'] || ''}));
  matchReview.slice(-8).forEach(row => items.push({type: 'Match Review', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Match Status'] || 'Needs Review', when: row['Last Updated'] || ''}));
  unmatchedConsents.slice(-8).forEach(row => items.push({type: 'Unmatched Consent', title: row['Child Full Name'] || row['Consent ResponseID'], detail: row['Best Prescreening Match'] || 'No eligible prescreening match', when: row['Submitted At'] || ''}));
  return items.sort((a, b) => String(b.when || '').localeCompare(String(a.when || ''))).slice(0, 15);
}

function buildReportData_(participants, followups, matchReview, unmatchedConsents, ready, screening, eligible, ineligible) {
  return {
    contactStatus: groupCounts_(followups, 'Follow Up Status'),
    consentStatus: groupCounts_(participants, 'consentStatus'),
    readiness: groupCounts_(participants, 'readyForDarts'),
    matchStatus: groupCounts_(participants, 'matchStatus'),
    reviewStatus: groupCounts_(participants, 'eligibilityReviewStatus'),
    humanEligibilityDecision: groupCounts_(participants, 'humanEligibilityDecision'),
    systemEligibilitySuggestion: groupCounts_(participants, 'systemEligibilitySuggestion'),
    actionItems: [
      {label: 'Eligibility pending review', value: (screening || []).filter(row => !row['Human Eligibility Decision'] || row['Human Eligibility Decision'] === 'Pending Review').length},
      {label: 'Needs more information', value: (screening || []).filter(row => row['Human Eligibility Decision'] === 'Needs More Information').length},
      {label: 'Approved eligible', value: (eligible || []).length},
      {label: 'Marked ineligible', value: (ineligible || []).length},
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
  const master = readObjectsOrEmpty_(ss, SHEETS.MASTER);
  const screening = readObjectsOrEmpty_(ss, SHEETS.SCREENING_REVIEW);
  const eligible = readObjectsOrEmpty_(ss, SHEETS.ELIGIBLE);
  const ineligible = readObjectsOrEmpty_(ss, SHEETS.INELIGIBLE);
  const prescreens = readObjectsOrEmpty_(ss, SHEETS.PRESCREEN_CLEAN);
  const consents = readObjectsOrEmpty_(ss, SHEETS.CONSENT_CLEAN);
  const ready = readObjectsOrEmpty_(ss, SHEETS.READY);
  const masterById = new Map(master.map(row => [String(row['EnrollmentID'] || ''), row]));
  const sourceRecords = screening.length ? screening : prescreens.map(row => screeningLikeFromPrescreen_(row));
  const participants = sourceRecords.map(row => {
    const masterRow = masterById.get(String(row['EnrollmentID'] || '')) || {};
    return participantFromScreening_(row, masterRow);
  });
  return {
    generatedAt: new Date().toISOString(),
    summary: buildProfessorSummary_(participants, prescreens, consents, ready, ss, master, eligible, ineligible),
    participants,
    screening,
    eligibleParticipants: eligible,
    ineligibleParticipants: ineligible,
    followups: participants.filter(p => p.manualContactNeeded === 'Yes' || p.humanEligibilityDecision === 'Needs More Information'),
    ready: participants.filter(p => p.readyForDarts === 'Yes'),
    needsReview: participants.filter(p => ['Pending Review', 'Needs More Information'].includes(p.humanEligibilityDecision) || p.eligibilityReviewStatus === 'Needs Review' || p.readyForDarts === 'Review')
  };
}

function participantFromScreening_(row, masterRow) {
  return {
    enrollmentId: row['EnrollmentID'] || masterRow['EnrollmentID'] || '',
    childName: row['Child Full Name'] || masterRow['Child Full Name'] || '',
    parentName: row['Parent/Caretaker Name'] || masterRow['Parent/Caretaker Name'] || '',
    parentEmail: row['Parent Email'] || masterRow['Parent Email'] || '',
    parentPhone: row['Parent Phone'] || masterRow['Parent Phone'] || '',
    cohortId: row['Cohort ID'] || masterRow['Cohort ID'] || '',
    cohortName: row['Cohort Name'] || masterRow['Cohort Name'] || '',
    site: row['Site'] || masterRow['Site'] || '',
    programTerm: row['Program Term'] || masterRow['Program Term'] || '',
    prescreeningStatus: 'Completed',
    consentStatus: masterRow['Consent Status'] || 'Pending',
    neurodivergentResponse: row['Neurodivergent Response'] || masterRow['Neurodivergent Response'] || '',
    conditions: row['Conditions/Diagnoses'] || '',
    supportDetails: row['Updated Support Details'] || row['Diagnostic/Support Details'] || '',
    originalSupportDetails: row['Diagnostic/Support Details'] || '',
    updatedSupportDetails: row['Updated Support Details'] || '',
    physicalSupports: row['Physical Disability Supports'] || '',
    systemEligibilitySuggestion: row['System Eligibility Suggestion'] || '',
    systemEligibilityReason: row['System Eligibility Reason'] || '',
    humanEligibilityDecision: row['Human Eligibility Decision'] || 'Pending Review',
    decisionReason: row['Decision Reason'] || '',
    manualContactNeeded: row['Manual Contact Needed'] || 'No',
    manualContactStatus: row['Manual Contact Status'] || 'Not Needed',
    followUpNeeded: row['Manual Contact Needed'] || masterRow['Follow Up Needed'] || 'No',
    followUpStatus: row['Manual Contact Status'] || masterRow['Follow Up Status'] || 'Not Needed',
    assignedTo: row['Assigned To'] || '',
    parentResponseSummary: row['Parent Response Summary'] || '',
    eligibilityReviewStatus: row['Eligibility Review Status'] || '',
    matchStatus: masterRow['Match Status'] || 'No Consent Yet',
    matchConfidenceScore: masterRow['Match Confidence Score'] || 0,
    matchReasons: masterRow['Match Reasons'] || '',
    needsMatchReview: masterRow['Needs Match Review'] || 'No',
    possibleConsentMatches: masterRow['Possible Consent Matches'] || '',
    reviewedBy: row['Reviewed By'] || '',
    reviewDate: row['Review Date'] || '',
    piNotes: row['PI Notes'] || '',
    notes: row['Notes'] || masterRow['Notes'] || '',
    enrollmentStatus: masterRow['Enrollment Status'] || row['Eligibility Review Status'] || '',
    readyForDarts: masterRow['Ready for DARTS'] || 'No',
    lastUpdated: row['Last Updated'] || masterRow['Last Updated'] || ''
  };
}


function screeningLikeFromPrescreen_(prescreen) {
  const suggestion = systemEligibilitySuggestion_(prescreen);
  return {
    'EnrollmentID': enrollmentIdFor_(prescreen),
    'PrescreeningID': prescreen['PrescreeningID'] || '',
    'ResponseID': prescreen['ResponseID'] || '',
    'Submitted At': prescreen['Submitted At'] || '',
    'Child Full Name': prescreen['Child Full Name'] || '',
    'Parent/Caretaker Name': prescreen['Parent/Caretaker Name'] || '',
    'Parent Email': prescreen['Parent Email'] || '',
    'Parent Phone': prescreen['Parent Phone'] || '',
    'Cohort ID': prescreen['Cohort ID'] || '',
    'Cohort Name': prescreen['Cohort Name'] || '',
    'Site': prescreen['Site'] || '',
    'Program Term': prescreen['Program Term'] || '',
    'Neurodivergent Response': prescreen['Neurodivergent Response'] || '',
    'Conditions/Diagnoses': prescreen['Conditions/Diagnoses'] || '',
    'Diagnostic/Support Details': prescreen['Diagnostic/Support Details'] || '',
    'Physical Disability Supports': prescreen['Physical Disability Supports'] || '',
    'System Eligibility Suggestion': suggestion,
    'System Eligibility Reason': systemEligibilityReason_(prescreen, suggestion),
    'Human Eligibility Decision': 'Pending Review',
    'Decision Reason': '',
    'Manual Contact Needed': contactNeededForDecision_('Pending Review', prescreen),
    'Manual Contact Status': 'Not Started',
    'Eligibility Review Status': screeningStatusForDecision_('Pending Review', suggestion),
    'Last Updated': prescreen['Submitted At'] || ''
  };
}


function buildProfessorSummary_(participants, prescreens, consents, ready, ss, master, eligible, ineligible) {
  const computed = {
    prescreened: prescreens.length,
    consentSubmitted: consents.length,
    consentCompleted: consents.filter(row => row['Consent Status'] === 'Completed').length,
    masterRecords: (master || []).length,
    neurodivergentYes: participants.filter(p => p.neurodivergentResponse === 'Yes').length,
    neurodivergentNo: participants.filter(p => p.neurodivergentResponse === 'No').length,
    eligibleApproved: (eligible || []).length,
    ineligible: (ineligible || []).length,
    followUpNeeded: participants.filter(p => p.manualContactNeeded === 'Yes' || p.followUpNeeded === 'Yes').length,
    followUpCompleted: participants.filter(p => p.manualContactStatus === 'Completed' || p.followUpStatus === 'Completed').length,
    readyForDarts: ready.length,
    needsReview: participants.filter(p => p.eligibilityReviewStatus === 'Needs Review' || p.readyForDarts === 'Review').length
  };
  return computed;
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
  const allowed = ['Follow Up Status', 'Follow Up Completed Date', 'Assigned To', 'Notes', ...FOLLOWUP_REVIEW_FIELDS, 'Human Eligibility Decision', 'Decision Reason', 'Manual Contact Needed', 'Manual Contact Status'];
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
    ['Screening Review Records', countRows_(ss, SHEETS.SCREENING_REVIEW)],
    ['Approved Eligible', countRows_(ss, SHEETS.ELIGIBLE)],
    ['Marked Ineligible', countRows_(ss, SHEETS.INELIGIBLE)],
    ['Total Master Enrollment Records', countRows_(ss, SHEETS.MASTER)],
    ['Manual Contact Needed', countWhere_(ss, SHEETS.SCREENING_REVIEW, 'Manual Contact Needed', 'Yes')],
    ['Manual Contact Completed', countWhere_(ss, SHEETS.SCREENING_REVIEW, 'Manual Contact Status', 'Completed')],
    ['Consent Completed', countWhere_(ss, SHEETS.MASTER, 'Consent Status', 'Completed')],
    ['Ready for DARTS', countWhere_(ss, SHEETS.MASTER, 'Ready for DARTS', 'Yes')],
    ['Unmatched Consent Records', countRows_(ss, UNMATCHED_CONSENT_SHEET)],
    ['Possible Duplicate Records', countWhere_(ss, MATCH_REVIEW_SHEET, 'Review Status', 'Needs Review')],
    ['Needs Review', countWhere_(ss, SHEETS.SCREENING_REVIEW, 'Eligibility Review Status', 'Needs Review')]
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
  const neuroRaw = getNeurodivergentRawResponse_(raw);
  const neuro = decodeQuestionProNeuroResponse_(neuroRaw);
  const conditions = collectOptionValues_(raw, ['Autism', 'ADHD', 'Dyslexia', 'Intellectual disability', 'Developmental disability', 'Epilepsy', 'Traumatic brain injury', 'OCD', 'Down Syndrome', 'Other genetic condition', 'Other (Please specify)']);
  const physicalSupports = getByContains_(raw, ['physically disabled', 'additional supports']);
  const followUpNeeded = needsFollowup_(neuro, conditions, physicalSupports, config) ? 'Yes' : 'No';
  return {
    'PrescreeningID': `PRE-${responseId}`,
    'ResponseID': responseId,
    'Submitted At': getFirst_(raw, ['Timestamp (mm/dd/yyyy)', 'Timestamp']),
    'Response Status': getFirst_(raw, ['Response Status']) || 'Completed',
    'Cohort ID': getFirst_(raw, ['Cohort ID', 'cohort_id', 'cohortId']),
    'Cohort Name': getFirst_(raw, ['Cohort Name', 'cohort_name', 'cohortName']),
    'Site': getFirst_(raw, ['Site', 'site']),
    'Program Term': getFirst_(raw, ['Program Term', 'program_term', 'term']),
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
    'Cohort ID': getFirst_(raw, ['Cohort ID', 'cohort_id', 'cohortId']),
    'Cohort Name': getFirst_(raw, ['Cohort Name', 'cohort_name', 'cohortName']),
    'Site': getFirst_(raw, ['Site', 'site']),
    'Program Term': getFirst_(raw, ['Program Term', 'program_term', 'term']),
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


function ensureScreeningInfrastructure_(ss) {
  ensureSheetWithHeaders_(ss, SHEETS.SCREENING_REVIEW, SCREENING_REVIEW_HEADERS);
  ensureSheetWithHeaders_(ss, SHEETS.ELIGIBLE, ELIGIBLE_PARTICIPANTS_HEADERS);
  ensureSheetWithHeaders_(ss, SHEETS.INELIGIBLE, INELIGIBLE_PARTICIPANTS_HEADERS);
}

function readManualScreeningState_(sheet) {
  const state = new Map();
  if (!sheet || sheet.getLastRow() < DATA_START_ROW) return state;
  readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW).forEach(row => {
    const id = String(row['EnrollmentID'] || '').trim();
    if (!id) return;
    state.set(id, row);
  });
  return state;
}

function buildScreeningReview_(prescreens, manualByEnrollment, legacyByEnrollment) {
  return prescreens.map(p => {
    const enrollmentId = enrollmentIdFor_(p);
    const manual = manualByEnrollment.get(enrollmentId) || {};
    const legacy = legacyByEnrollment.get(enrollmentId) || {};
    const suggestion = systemEligibilitySuggestion_(p);
    const decision = manual['Human Eligibility Decision'] || legacyDecisionFromFollowup_(legacy) || 'Pending Review';
    const reason = manual['Decision Reason'] || legacy['Follow-Up Outcome'] || '';
    const manualContactNeeded = manual['Manual Contact Needed'] || contactNeededForDecision_(decision, p);
    const reviewStatus = manual['Eligibility Review Status'] || screeningStatusForDecision_(decision, suggestion);
    return {
      'EnrollmentID': enrollmentId,
      'PrescreeningID': p['PrescreeningID'] || '',
      'ResponseID': p['ResponseID'] || '',
      'Submitted At': p['Submitted At'] || '',
      'Child Full Name': manual['Child Full Name'] || p['Child Full Name'] || '',
      'Parent/Caretaker Name': manual['Parent/Caretaker Name'] || p['Parent/Caretaker Name'] || '',
      'Parent Email': manual['Parent Email'] || p['Parent Email'] || '',
      'Parent Phone': manual['Parent Phone'] || p['Parent Phone'] || '',
      'Cohort ID': manual['Cohort ID'] || p['Cohort ID'] || '',
      'Cohort Name': manual['Cohort Name'] || p['Cohort Name'] || '',
      'Site': manual['Site'] || p['Site'] || '',
      'Program Term': manual['Program Term'] || p['Program Term'] || '',
      'Neurodivergent Response': p['Neurodivergent Response'] || '',
      'Conditions/Diagnoses': p['Conditions/Diagnoses'] || '',
      'Diagnostic/Support Details': p['Diagnostic/Support Details'] || '',
      'Physical Disability Supports': p['Physical Disability Supports'] || '',
      'System Eligibility Suggestion': suggestion,
      'System Eligibility Reason': systemEligibilityReason_(p, suggestion),
      'Human Eligibility Decision': decision,
      'Decision Reason': reason,
      'Manual Contact Needed': manualContactNeeded,
      'Manual Contact Status': manual['Manual Contact Status'] || legacy['Follow Up Status'] || (manualContactNeeded === 'Yes' ? 'Not Started' : 'Not Needed'),
      'Assigned To': manual['Assigned To'] || legacy['Assigned To'] || '',
      'Updated Support Details': manual['Updated Support Details'] || legacy['Updated Support Details'] || '',
      'Parent Response Summary': manual['Parent Response Summary'] || legacy['Parent Response Summary'] || legacy['Follow-Up Outcome'] || '',
      'Eligibility Review Status': reviewStatus,
      'Reviewed By': manual['Reviewed By'] || legacy['Reviewed By'] || '',
      'Review Date': manual['Review Date'] || legacy['Review Date'] || '',
      'PI Notes': manual['PI Notes'] || legacy['PI Notes'] || '',
      'Notes': manual['Notes'] || legacy['Notes'] || '',
      'Last Updated': new Date()
    };
  });
}

function systemEligibilitySuggestion_(p) {
  const neuro = String(p['Neurodivergent Response'] || '').trim();
  if (neuro === 'Yes') return 'Likely Eligible';
  if (neuro === 'No') return hasSupportContext_(p) ? 'Needs Human Review' : 'Likely Not Eligible';
  return 'Needs Human Review';
}

function systemEligibilityReason_(p, suggestion) {
  if (suggestion === 'Likely Eligible') return 'Parent selected Yes for neurodivergent/disability/developmental condition/learning difference.';
  if (suggestion === 'Likely Not Eligible') return 'Parent selected No and no additional diagnostic/support details were detected.';
  if (String(p['Neurodivergent Response'] || '') === 'No') return 'Parent selected No, but support details or conditions were detected; staff should review.';
  return 'Neurodivergent/support-needs answer is missing or ambiguous; staff should review.';
}

function hasSupportContext_(p) {
  return Boolean(String(p['Conditions/Diagnoses'] || p['Diagnostic/Support Details'] || p['Physical Disability Supports'] || '').trim());
}

function legacyDecisionFromFollowup_(legacy) {
  const status = String(legacy['Eligibility Review Status'] || '').trim();
  if (/not eligible/i.test(status)) return 'Not Eligible';
  if (/ready|likely eligible/i.test(status)) return 'Approved Eligible';
  return '';
}

function contactNeededForDecision_(decision, p) {
  if (decision === 'Needs More Information') return 'Yes';
  if (String(p['Neurodivergent Response'] || '') === 'No' && hasSupportContext_(p)) return 'Yes';
  return 'No';
}

function screeningStatusForDecision_(decision, suggestion) {
  if (decision === 'Approved Eligible') return 'Eligible';
  if (decision === 'Not Eligible') return 'Ineligible';
  if (decision === 'Needs More Information') return 'Needs More Information';
  return suggestion === 'Needs Human Review' ? 'Needs Review' : 'Pending Review';
}

function buildEligibleParticipants_(screeningRows) {
  return screeningRows.filter(row => row['Human Eligibility Decision'] === 'Approved Eligible').map(row => Object.assign({}, row, {
    'Eligibility Approved At': row['Review Date'] || row['Last Updated'] || new Date()
  }));
}

function buildIneligibleParticipants_(screeningRows) {
  return screeningRows.filter(row => row['Human Eligibility Decision'] === 'Not Eligible').map(row => Object.assign({}, row, {
    'Eligibility Closed At': row['Review Date'] || row['Last Updated'] || new Date()
  }));
}

function buildManualContactQueue_(screeningRows) {
  return screeningRows.filter(row => row['Manual Contact Needed'] === 'Yes' || row['Human Eligibility Decision'] === 'Needs More Information').map(row => ({
    'EnrollmentID': row['EnrollmentID'],
    'Child Full Name': row['Child Full Name'],
    'Parent/Caretaker Name': row['Parent/Caretaker Name'],
    'Parent Email': row['Parent Email'],
    'Parent Phone': row['Parent Phone'],
    'Cohort ID': row['Cohort ID'],
    'Cohort Name': row['Cohort Name'],
    'Site': row['Site'],
    'Program Term': row['Program Term'],
    'Neurodivergent Response': row['Neurodivergent Response'],
    'Conditions/Diagnoses': row['Conditions/Diagnoses'],
    'Support Details': row['Updated Support Details'] || row['Diagnostic/Support Details'],
    'Physical Disability Supports': row['Physical Disability Supports'],
    'Follow Up Reason': row['System Eligibility Reason'],
    'Follow Up Status': row['Manual Contact Status'] || 'Not Started',
    'Assigned To': row['Assigned To'],
    'Notes': row['Notes'],
    'Updated Support Details': row['Updated Support Details'],
    'Parent Response Summary': row['Parent Response Summary'],
    'Eligibility Review Status': row['Eligibility Review Status'],
    'Human Eligibility Decision': row['Human Eligibility Decision'],
    'Decision Reason': row['Decision Reason'],
    'Reviewed By': row['Reviewed By'],
    'Review Date': row['Review Date'],
    'PI Notes': row['PI Notes']
  }));
}


function updateParticipantCoreData(updates) {
  updates = updates || {};
  if (!updates.enrollmentId) throw new Error('Missing enrollmentId.');
  const ss = getSpreadsheet_();
  const screeningSheet = ss.getSheetByName(SHEETS.SCREENING_REVIEW);
  const screeningRows = readObjects_(screeningSheet, CLEAN_HEADER_ROW, DATA_START_ROW);
  const current = screeningRows.find(row => String(row['EnrollmentID']) === String(updates.enrollmentId));
  if (!current) throw new Error(`EnrollmentID not found in Screening_Review: ${updates.enrollmentId}`);
  const responseId = String(current['ResponseID'] || '').trim();
  if (!responseId) throw new Error('This participant does not have a prescreening ResponseID to update.');

  const screeningHeaders = screeningSheet.getRange(1, 1, 1, screeningSheet.getLastColumn()).getValues()[0].map(String);
  const screeningRow = screeningRows.findIndex(row => String(row['EnrollmentID']) === String(updates.enrollmentId)) + DATA_START_ROW;
  const screeningUpdates = {
    'Child Full Name': updates.childName,
    'Parent/Caretaker Name': updates.parentName,
    'Parent Email': updates.parentEmail,
    'Parent Phone': updates.parentPhone,
    'Cohort ID': updates.cohortId,
    'Cohort Name': updates.cohortName,
    'Site': updates.site,
    'Program Term': updates.programTerm
  };
  Object.entries(screeningUpdates).forEach(([field, value]) => {
    if (value === undefined) return;
    const col = screeningHeaders.indexOf(field) + 1;
    if (col > 0) screeningSheet.getRange(screeningRow, col).setValue(value);
  });

  const rawSheet = ss.getSheetByName(SHEETS.PRESCREEN_RAW);
  const rawHeaders = rawSheet.getRange(RAW_HEADER_ROW, 1, 1, rawSheet.getLastColumn()).getValues()[0].map(String);
  const rawValues = rawSheet.getDataRange().getValues();
  const responseCol = rawHeaders.findIndex(h => /response\s*id/i.test(String(h))) + 1;
  if (responseCol < 1) throw new Error('Prescreening_Raw is missing a Response ID column.');
  const rawRowIndex = rawValues.findIndex((row, index) => index >= RAW_DATA_START_ROW - 1 && String(row[responseCol - 1]) === responseId);
  if (rawRowIndex < RAW_DATA_START_ROW - 1) throw new Error(`Could not find raw prescreening response ${responseId}.`);
  const targetRow = rawRowIndex + 1;

  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Childs Name (First Last)', 'Child Full Name', 'Child Name'], updates.childName);
  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Parent / Caretaker Name (First Last):', 'Parent/Caretaker Name', 'Parent Name'], updates.parentName);
  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Email Address:', 'Parent Email', 'Respondent Email'], updates.parentEmail);
  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Primary Phone Number. Note that your phone number should be 10-digits and in this format. XXX-XXX-XXXX', 'Parent Phone', 'Phone'], updates.parentPhone);
  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Cohort ID'], updates.cohortId);
  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Cohort Name'], updates.cohortName);
  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Site'], updates.site);
  setRawValueForAnyHeader_(rawSheet, rawHeaders, targetRow, ['Program Term'], updates.programTerm);

  refreshEnrollmentTracker();
  return buildAppDashboardData_(false);
}

function setRawValueForAnyHeader_(sheet, headers, row, possibleHeaders, value) {
  if (value === undefined) return;
  const normalized = possibleHeaders.map(normalizeHeaderKey_);
  const index = headers.findIndex(h => normalized.includes(normalizeHeaderKey_(h)));
  if (index >= 0) sheet.getRange(row, index + 1).setValue(value);
}

function updateScreeningReview(updates) {
  updates = updates || {};
  if (!updates.enrollmentId) throw new Error('Missing enrollmentId.');
  const ss = getSpreadsheet_();
  ensureScreeningInfrastructure_(ss);
  const sheet = ss.getSheetByName(SHEETS.SCREENING_REVIEW);
  const rows = readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW);
  const rowIndex = rows.findIndex(row => String(row['EnrollmentID']) === String(updates.enrollmentId));
  if (rowIndex < 0) throw new Error(`EnrollmentID not found in Screening_Review: ${updates.enrollmentId}`);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const targetRow = rowIndex + DATA_START_ROW;
  SCREENING_MANUAL_FIELDS.forEach(field => {
    if (updates[field] === undefined) return;
    const col = headers.indexOf(field) + 1;
    if (col > 0) sheet.getRange(targetRow, col).setValue(updates[field]);
  });
  if (updates['Human Eligibility Decision'] && !updates['Review Date']) {
    const col = headers.indexOf('Review Date') + 1;
    if (col > 0) sheet.getRange(targetRow, col).setValue(new Date());
  }
  refreshEnrollmentTracker();
  return buildAppDashboardData_(false);
}

function batchUpdateScreeningReview(request) {
  request = request || {};
  const ids = Array.isArray(request.enrollmentIds) ? request.enrollmentIds : [];
  if (!ids.length) throw new Error('Select at least one participant.');
  const decision = request.decision || 'Approved Eligible';
  const ss = getSpreadsheet_();
  ensureScreeningInfrastructure_(ss);
  const sheet = ss.getSheetByName(SHEETS.SCREENING_REVIEW);
  const rows = readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idSet = new Set(ids.map(String));
  rows.forEach((row, i) => {
    if (!idSet.has(String(row['EnrollmentID']))) return;
    const targetRow = i + DATA_START_ROW;
    const updates = {
      'Human Eligibility Decision': decision,
      'Decision Reason': request.reason || '',
      'Eligibility Review Status': screeningStatusForDecision_(decision, { 'Approved Eligible':'Likely Eligible' }[decision] || ''),
      'Reviewed By': request.reviewedBy || '',
      'Review Date': new Date(),
      'Notes': request.notes || row['Notes'] || ''
    };
    Object.entries(updates).forEach(([field, value]) => {
      const col = headers.indexOf(field) + 1;
      if (col > 0) sheet.getRange(targetRow, col).setValue(value);
    });
  });
  refreshEnrollmentTracker();
  return buildAppDashboardData_(false);
}

function buildMasterEnrollment_(eligibleRows, consents, config, manualMatchByEnrollment) {
  return eligibleRows.map(p => {
    const enrollmentId = p['EnrollmentID'] || enrollmentIdFor_(p);
    const manualMatch = manualMatchByEnrollment.get(enrollmentId) || {};
    const match = selectConsentMatch_(p, consents, manualMatch);
    const consent = match.accepted ? match.consent : null;
    const consentStatus = consent ? consent['Consent Status'] : (match.needsReview ? 'Review' : config.defaultConsentStatus);
    const ready = isReadyWithEligibility_(p['Human Eligibility Decision'], consentStatus, match);
    const enrollmentStatus = ready === 'Yes' ? 'Ready for DARTS' : (match.needsReview ? 'Ready for Review' : 'Eligible - Awaiting Consent');
    return {
      'EnrollmentID': enrollmentId,
      'Child Full Name': p['Child Full Name'],
      'Parent/Caretaker Name': p['Parent/Caretaker Name'],
      'Parent Email': p['Parent Email'] || (consent && consent['Parent Email']) || '',
      'Parent Phone': p['Parent Phone'] || (consent && consent['Parent Phone']) || '',
      'Cohort ID': p['Cohort ID'] || (consent && consent['Cohort ID']) || '',
      'Cohort Name': p['Cohort Name'] || (consent && consent['Cohort Name']) || '',
      'Site': p['Site'] || (consent && consent['Site']) || '',
      'Program Term': p['Program Term'] || (consent && consent['Program Term']) || '',
      'Human Eligibility Decision': p['Human Eligibility Decision'],
      'Eligibility Review Status': p['Eligibility Review Status'] || 'Eligible',
      'Decision Reason': p['Decision Reason'] || '',
      'Prescreening Status': 'Completed',
      'Consent Status': consentStatus,
      'Neurodivergent Response': p['Neurodivergent Response'],
      'Follow Up Needed': p['Manual Contact Needed'] || 'No',
      'Follow Up Status': p['Manual Contact Status'] || 'Not Needed',
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
      'Notes': p['Notes'] || ''
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
    'Cohort ID': m['Cohort ID'],
    'Cohort Name': m['Cohort Name'],
    'Site': m['Site'],
    'Program Term': m['Program Term'],
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
      'Cohort ID': c['Cohort ID'],
      'Cohort Name': c['Cohort Name'],
      'Site': c['Site'],
      'Program Term': c['Program Term'],
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
  const prescreenById = new Map(prescreens.map(p => [p['EnrollmentID'] || enrollmentIdFor_(p), p]));
  return master.filter(m => m['Follow Up Needed'] === 'Yes').map(m => {
    const p = prescreenById.get(m['EnrollmentID']) || {};
    const manual = manualFollowup.get(m['EnrollmentID']) || {};
    return {
      'EnrollmentID': m['EnrollmentID'],
      'Child Full Name': m['Child Full Name'],
      'Parent/Caretaker Name': m['Parent/Caretaker Name'],
      'Parent Email': m['Parent Email'],
      'Parent Phone': m['Parent Phone'],
      'Cohort ID': m['Cohort ID'],
      'Cohort Name': m['Cohort Name'],
      'Site': m['Site'],
      'Program Term': m['Program Term'],
      'Neurodivergent Response': m['Neurodivergent Response'],
      'Conditions/Diagnoses': p['Conditions/Diagnoses'],
      'Support Details': p['Diagnostic/Support Details'],
      'Physical Disability Supports': p['Physical Disability Supports'],
      'Follow Up Reason': followupReason_(p),
      'Follow Up Status': manual['Follow Up Status'] || config.defaultFollowUpStatus,
      'Follow Up Completed Date': manual['Follow Up Completed Date'] || '',
      'Assigned To': manual['Assigned To'] || '',
      'Notes': manual['Notes'] || '',
      'Updated Support Details': manual['Updated Support Details'] || '',
      'Parent Response Summary': manual['Parent Response Summary'] || '',
      'Eligibility Review Status': manual['Eligibility Review Status'] || '',
      'Human Eligibility Decision': manual['Human Eligibility Decision'] || '',
      'Decision Reason': manual['Decision Reason'] || '',
      'Reviewed By': manual['Reviewed By'] || '',
      'Review Date': manual['Review Date'] || '',
      'PI Notes': manual['PI Notes'] || ''
    };
  });
}


function buildReadyForDarts_(master, prescreens, consents) {
  const prescreenById = new Map(prescreens.map(p => [p['EnrollmentID'] || enrollmentIdFor_(p), p]));
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
      'Cohort ID': m['Cohort ID'],
      'Cohort Name': m['Cohort Name'],
      'Site': m['Site'],
      'Program Term': m['Program Term'],
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

function installRawTabAutomation() {
  const ss = getSpreadsheet_();
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction && trigger.getHandlerFunction() === 'handleWorkbookChange')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('handleWorkbookChange').forSpreadsheet(ss).onChange().create();
  rememberRawTabSignatures_(ss);
  return getTrackerSummary();
}

function handleWorkbookChange(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;
  try {
    const ss = getSpreadsheet_();
    if (rawTabsChanged_(ss)) refreshEnrollmentTracker();
  } finally {
    lock.releaseLock();
  }
}

function rawTabsChanged_(ss) {
  const props = PropertiesService.getScriptProperties();
  const current = rawTabSignatures_(ss);
  const previous = props.getProperty('RAW_TAB_SIGNATURES') || '';
  return current !== previous;
}

function rememberRawTabSignatures_(ss) {
  PropertiesService.getScriptProperties().setProperty('RAW_TAB_SIGNATURES', rawTabSignatures_(ss));
}

function rawTabSignatures_(ss) {
  return [SHEETS.PRESCREEN_RAW, SHEETS.CONSENT_RAW].map(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return `${sheetName}:missing`;
    const rows = readRawRows_(sheet);
    const ids = rows.map(row => String(row['Response ID'] || row['ResponseID'] || row['Timestamp (mm/dd/yyyy)'] || row['Timestamp'] || '')).join('|');
    return `${sheetName}:${rows.length}:${ids}`;
  }).join('||');
}


function ensureCohortInfrastructure_(ss) {
  ensureRawCohortHeaders_(ss);
  ensureSheetWithHeaders_(ss, COHORTS_SHEET, COHORT_HEADERS);
}

function ensureRawCohortHeaders_(ss) {
  [SHEETS.PRESCREEN_RAW, SHEETS.CONSENT_RAW].forEach(name => appendMissingHeadersAtRow_(ss.getSheetByName(name), COHORT_FIELDS, RAW_HEADER_ROW));
}

function ensureCoreOutputHeaders_(ss) {
  [SHEETS.PRESCREEN_CLEAN, SHEETS.CONSENT_CLEAN, SHEETS.SCREENING_REVIEW, SHEETS.ELIGIBLE, SHEETS.INELIGIBLE, SHEETS.MASTER, SHEETS.FOLLOWUP, SHEETS.READY, MATCH_REVIEW_SHEET, UNMATCHED_CONSENT_SHEET]
    .forEach(name => appendMissingHeaders_(ss.getSheetByName(name), COHORT_FIELDS));
}

function getRequestedCohort_(e, payload) {
  const source = Object.assign({}, payload || {}, (e && e.parameter) || {});
  return {
    cohortId: String(source.cohort_id || source.cohortId || source['Cohort ID'] || source.cohort || '').trim(),
    cohortName: String(source.cohort_name || source.cohortName || source['Cohort Name'] || '').trim(),
    site: String(source.site || source.Site || '').trim(),
    programTerm: String(source.program_term || source.programTerm || source.term || source['Program Term'] || '').trim()
  };
}

function cohortFieldsFrom_(cohort) {
  cohort = cohort || {};
  return {
    'Cohort ID': cohort.cohortId || '',
    'Cohort Name': cohort.cohortName || '',
    'Site': cohort.site || '',
    'Program Term': cohort.programTerm || ''
  };
}

function upsertCohortFromWebhook_(cohort) {
  if (!cohort || !cohort.cohortId) return;
  const ss = getSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(ss, COHORTS_SHEET, COHORT_HEADERS);
  const rows = readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW);
  const targetIndex = rows.findIndex(row => String(row['Cohort ID'] || '').toLowerCase() === cohort.cohortId.toLowerCase());
  const rowNumber = targetIndex >= 0 ? targetIndex + DATA_START_ROW : sheet.getLastRow() + 1;
  const valuesByHeader = {
    'Cohort ID': cohort.cohortId,
    'Cohort Name': cohort.cohortName || cohort.cohortId,
    'Site': cohort.site || '',
    'Program Term': cohort.programTerm || '',
    'Status': targetIndex >= 0 ? rows[targetIndex]['Status'] || 'Active' : 'Active',
    'Prescreening Webhook URL': buildWebhookUrl_('prescreening', cohort),
    'Notes': targetIndex >= 0 ? rows[targetIndex]['Notes'] || '' : '',
    'Last Seen': new Date()
  };
  COHORT_HEADERS.forEach((header, i) => sheet.getRange(rowNumber, i + 1).setValue(valuesByHeader[header] || ''));
}

function buildCohortSummaries_(ss, participants, prescreens, consents) {
  const configured = readObjectsOrEmpty_(ss, COHORTS_SHEET);
  const byId = new Map();
  configured.forEach(row => {
    const id = String(row['Cohort ID'] || '').trim();
    if (!id) return;
    byId.set(id, {
      cohortId: id,
      cohortName: row['Cohort Name'] || id,
      site: row['Site'] || '',
      programTerm: row['Program Term'] || '',
      status: row['Status'] || '',
      prescreened: 0,
      consentSubmitted: 0,
      masterRecords: 0,
      followUpNeeded: 0,
      readyForDarts: 0
    });
  });
  function ensure(id, fallback) {
    id = String(id || 'Unassigned').trim() || 'Unassigned';
    if (!byId.has(id)) byId.set(id, Object.assign({cohortId: id, cohortName: id, site: '', programTerm: '', status: id === 'Unassigned' ? 'Needs cohort assignment' : '', prescreened: 0, consentSubmitted: 0, masterRecords: 0, followUpNeeded: 0, readyForDarts: 0}, fallback || {}));
    return byId.get(id);
  }
  prescreens.forEach(row => ensure(row['Cohort ID'], {cohortName: row['Cohort Name'] || row['Cohort ID'] || 'Unassigned', site: row['Site'] || '', programTerm: row['Program Term'] || ''}).prescreened++);
  participants.forEach(p => {
    const cohort = ensure(p.cohortId, {cohortName: p.cohortName || p.cohortId || 'Unassigned', site: p.site || '', programTerm: p.programTerm || ''});
    cohort.masterRecords++;
    if (p.consentStatus && p.consentStatus !== 'Pending' && p.consentStatus !== 'Review') cohort.consentSubmitted++;
    if (p.followUpNeeded === 'Yes') cohort.followUpNeeded++;
    if (p.readyForDarts === 'Yes') cohort.readyForDarts++;
  });
  consents.filter(row => row['Cohort ID']).forEach(row => ensure(row['Cohort ID'], {cohortName: row['Cohort Name'] || row['Cohort ID'], site: row['Site'] || '', programTerm: row['Program Term'] || ''}));
  return Array.from(byId.values()).sort((a, b) => String(a.cohortId).localeCompare(String(b.cohortId)));
}

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
  if (![SHEETS.PRESCREEN_CLEAN, SHEETS.SCREENING_REVIEW, SHEETS.ELIGIBLE, SHEETS.INELIGIBLE, SHEETS.FOLLOWUP, SHEETS.MASTER, SHEETS.READY].includes(sheet.getName())) return;
  const rowsToFormat = Math.max(dataRows, Math.max(0, sheet.getLastRow() - 1), 1);
  const range = sheet.getRange(2, 1, rowsToFormat, columns);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment('middle');
  sheet.setRowHeights(2, rowsToFormat, 24);
}

function applyRawSheetFormatting_(sheet) {
  if (!sheet) return;
  const rows = Math.max(sheet.getMaxRows(), sheet.getLastRow(), RAW_DATA_START_ROW);
  const columns = Math.max(sheet.getMaxColumns(), sheet.getLastColumn(), 1);
  sheet.getRange(1, 1, rows, columns)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');
  if (rows >= RAW_DATA_START_ROW) sheet.setRowHeights(RAW_DATA_START_ROW, rows - RAW_DATA_START_ROW + 1, 24);
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

function appendMissingHeadersAtRow_(sheet, fields, headerRow) {
  if (!sheet) return;
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0].map(String);
  fields.forEach(field => {
    if (headers.includes(field)) return;
    sheet.getRange(headerRow, sheet.getLastColumn() + 1).setValue(field);
    headers.push(field);
  });
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
  // Study-team rule: any score of 60 or higher is accepted automatically.
  // Lower-confidence matches remain visible in Match_Review for manual approval.
  const accepted = best.score >= 60;
  const needsReview = !accepted && best.score >= 40;
  return {
    consent: accepted ? best.consent : null,
    accepted,
    status: accepted ? 'Matched' : (needsReview ? 'Needs Review' : 'No Consent Yet'),
    score: best.score,
    reasons: closeSecond && !accepted ? best.reasons.concat(['Another consent record has a similar score; manual review recommended']) : best.reasons,
    needsReview: needsReview || (!accepted && Boolean(closeSecond)),
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
  if (prescreen['Cohort ID'] && consent['Cohort ID'] && String(prescreen['Cohort ID']).toLowerCase() === String(consent['Cohort ID']).toLowerCase()) { score += 8; reasons.push('Cohort exact'); }
  if (prescreen['Site'] && consent['Site'] && String(prescreen['Site']).toLowerCase() === String(consent['Site']).toLowerCase()) { score += 3; reasons.push('Site exact'); }
  if (prescreen['Cohort ID'] && consent['Cohort ID'] && String(prescreen['Cohort ID']).toLowerCase() !== String(consent['Cohort ID']).toLowerCase()) { score -= 10; reasons.push('Different cohort values; review recommended'); }
  return {score: Math.max(0, Math.min(score, 100)), reasons};
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
function isReadyWithEligibility_(decision, consentStatus, match) { return decision === 'Approved Eligible' && match.accepted && consentStatus === 'Completed' ? 'Yes' : (match.needsReview ? 'Review' : 'No'); }

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

function getNeurodivergentRawResponse_(raw) {
  const direct = getByHeaderPattern_(raw, isNeurodivergentQuestionKey_);
  if (direct !== '') return direct;

  const fallbackPayloads = ['External Reference', 'Custom Variable 5', '_rawPayload', 'complete_response', 'completeResponse', 'rawBody', 'payload']
    .map(key => raw[key])
    .filter(Boolean);
  for (const fallbackPayload of fallbackPayloads) {
    const parsed = parseJsonIfPossible_(fallbackPayload);
    if (!parsed) continue;
    const extracted = extractQuestionProResponseSet_(parsed);
    const extractedDirect = getByHeaderPattern_(extracted, isNeurodivergentQuestionKey_);
    if (extractedDirect !== '') return extractedDirect;
  }

  return getByContains_(raw, ['identify as neurodivergent', 'diagnosed disability']);
}

function isNeurodivergentQuestionKey_(key) {
  const normalized = normalizeHeaderKey_(key);
  return normalized.includes('doesyourchildidentifyasneurodivergent') ||
    normalized.includes('neurodivergentresponse') ||
    (normalized.includes('neurodivergent') && !normalized.includes('neurodiverselearners')) ||
    (normalized.includes('diagnoseddisability') && normalized.includes('developmental')) ||
    (normalized.includes('learningdifference') && normalized.includes('disability'));
}

function getByHeaderPattern_(obj, predicate) {
  for (const [key, value] of Object.entries(obj || {})) {
    if (value !== '' && value !== null && value !== undefined && !isRawPayloadValue_(value) && predicate(key)) return value;
  }
  return '';
}

function parseJsonIfPossible_(value) {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  if (!/^[\[{]/.test(trimmed)) return null;
  try { return JSON.parse(trimmed); } catch (err) { return null; }
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
  // Study workflow note: Phase 1 follow-up is for families who answered
  // "No" to the neurodivergent/disability question so the team can ask for
  // clarification about support needs, diagnoses, accommodations, or services.
  // Families who answered "Yes" already gave the expected eligibility signal
  // and should not be automatically placed in Followup_Queue.
  return String(neuro || '').trim().toLowerCase() === 'no';
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
  if (!contents) return expandEmbeddedJsonPayloads_(payload);
  try {
    const parsed = JSON.parse(contents);
    return expandEmbeddedJsonPayloads_(Object.assign(payload, parsed));
  } catch (err) {
    // QuestionPro may send form-encoded bodies instead of raw JSON.
    if (contents.includes('=')) Object.assign(payload, parseFormEncoded_(contents));
    else payload.rawBody = contents;
    return expandEmbeddedJsonPayloads_(payload);
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
    const separator = pair.indexOf('=');
    const rawKey = separator >= 0 ? pair.slice(0, separator) : pair;
    const rawValue = separator >= 0 ? pair.slice(separator + 1) : '';
    const key = decodeURIComponent(rawKey.replace(/\+/g, ' '));
    obj[key] = decodeURIComponent(rawValue.replace(/\+/g, ' '));
    return obj;
  }, {});
}
function appendRawPayload_(formType, payload, cohort) {
  ensureRawCohortHeaders_(getSpreadsheet_());
  const flattened = flattenPayload_(payload);
  Object.assign(flattened, cohortFieldsFrom_(cohort));
  Object.assign(flattened, extractQuestionProResponseSet_(payload));
  removeRoutingOnlyFields_(flattened);
  flattened._rawPayload = JSON.stringify(payload);
  appendRawObjectForTest_(formType === 'Consent' ? SHEETS.CONSENT_RAW : SHEETS.PRESCREEN_RAW, flattened);
}

function removeRoutingOnlyFields_(obj) {
  ['form', 'formType', 'surveyType', 'cohort_id', 'cohort_name', 'cohortId', 'cohortName', 'program_term'].forEach(key => delete obj[key]);
}
function expandEmbeddedJsonPayloads_(payload) {
  const expanded = Object.assign({}, payload);
  ['complete_response', 'completeResponse', 'response', 'payload', 'rawBody'].forEach(key => {
    const value = expanded[key];
    if (typeof value !== 'string') return;
    const trimmed = value.trim();
    if (!trimmed || !/^[\[{]/.test(trimmed)) return;
    try {
      const parsed = JSON.parse(trimmed);
      Object.assign(expanded, parsed);
      expanded[key] = value;
    } catch (err) {
      // Keep the original field for audit purposes when it is not valid JSON.
    }
  });
  return expanded;
}

function extractQuestionProResponseSet_(payload) {
  const out = {};
  findQuestionProResponseSets_(payload).forEach(responseSet => {
    responseSet.forEach(item => {
      const key = item.questionText || item.questionDescription || item.questionCode || item.questionID;
      const value = extractQuestionProAnswer_(item);
      if (key && value !== '') out[key] = value;
      if (item.questionCode && value !== '') out[item.questionCode] = value;
    });
  });
  return out;
}

function findQuestionProResponseSets_(value, sets = []) {
  if (!value || typeof value !== 'object') return sets;
  if (Array.isArray(value)) {
    if (value.some(item => item && typeof item === 'object' && (item.questionText || item.questionDescription || item.questionCode || item.questionID))) {
      sets.push(value);
    } else {
      value.forEach(item => findQuestionProResponseSets_(item, sets));
    }
    return sets;
  }
  Object.entries(value).forEach(([key, child]) => {
    if (key === 'responseSet' && Array.isArray(child)) findQuestionProResponseSets_(child, sets);
    else findQuestionProResponseSets_(child, sets);
  });
  return sets;
}
function extractQuestionProAnswer_(item) {
  const directFields = ['answerText', 'answer', 'responseText', 'response', 'value', 'answerValue', 'selectedAnswer', 'displayText', 'answerCode', 'answerID', 'optionCode', 'recodeValue'];
  for (const field of directFields) {
    const value = normalizeCellValue_(item[field]);
    if (value !== '') return value;
  }
  const answerContainerFields = ['answers', 'answerValues', 'values', 'selectedAnswers', 'responseAnswers', 'answerOptions', 'selectedOptions'];
  for (const field of answerContainerFields) {
    if (Array.isArray(item[field])) {
      const values = item[field].map(normalizeCellValue_).filter(Boolean);
      if (values.length) return values.join('; ');
    } else if (item[field] && typeof item[field] === 'object') {
      const value = normalizeCellValue_(item[field]);
      if (value !== '') return value;
    }
  }
  return '';
}
function normalizeCellValue_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value;
  if (Array.isArray(value)) return value.map(normalizeCellValue_).filter(Boolean).join('; ');
  if (typeof value !== 'object') return String(value).trim();
  const answerFields = ['answerText', 'text', 'value', 'answerValue', 'label', 'displayText', 'optionText', 'option', 'code', 'answerCode', 'optionCode', 'recodeValue', 'name', 'email', 'phone'];
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
  applyRawSheetFormatting_(sh);
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
  setRawFallback_(headers, row, [
    'Does your child identify as neurodivergent or have a diagnosed disability developmental condition or learning difference?',
    'Does your child identify as neurodivergent or have a diagnosed disability developmental',
    'Neurodivergent Response'
  ], getNeurodivergentRawResponse_(obj));
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
function followupReason_(p) { return p['Neurodivergent Response'] === 'No' ? 'Neurodivergent/disability response was No; follow up for clarification about support needs or eligibility context' : ['Conditions/diagnoses listed', p['Physical Disability Supports'] && 'Physical support needs'].filter(Boolean).join('; '); }
function cleanName_(v) { return String(v || '').trim().replace(/\s+/g, ' ').replace(/\b\w/g, c => c.toUpperCase()); }
function normalizeEmail_(v) { return String(v || '').trim().toLowerCase(); }
function normalizePhone_(v) { const d = String(v || '').replace(/\D/g, ''); return d.length === 10 ? `${d.slice(0,3)}-${d.slice(3,6)}-${d.slice(6)}` : String(v || '').trim(); }
function normalizeToken_(v) { return String(v || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function countRows_(ss, name) { const sheet = ss.getSheetByName(name); return sheet ? readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW).length : 0; }
function countWhere_(ss, name, field, expected) { const sheet = ss.getSheetByName(name); return sheet ? readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW).filter(r => r[field] === expected).length : 0; }
function jsonSafe_(obj) { return JSON.parse(JSON.stringify(obj || {})); }
function jsonResponse_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function ensureRequiredSheets_(ss) { Object.values(SHEETS).forEach(n => { if (!ss.getSheetByName(n)) throw new Error(`Missing required sheet: ${n}`); }); }
