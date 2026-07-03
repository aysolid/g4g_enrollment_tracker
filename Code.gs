/**
 * Gaming4Good Enrollment Tracker — Phase 1 Apps Script automation.
 *
 * Workflow:
 * QuestionPro webhook/manual raw import -> Prescreening_Raw/Consent_Raw audit tabs
 * -> normalized clean tabs -> Master_Enrollment -> Followup_Queue and Ready_For_DARTS.
 *
 * Email automation is supervised: Apps Script drafts and approves email jobs, while
 * Power Automate + Outlook handles actual sending/reply callbacks once configured.
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
  'Reviewed By', 'Review Date', 'PI Notes',
  'Email Workflow Status', 'Latest EmailJobID', 'Latest Email Type', 'Latest Email Subject',
  'Latest Email Approved At', 'Latest Email Sent At', 'Latest Reply Received At',
  'Reply Status', 'Reminder Status', 'Reminder Count', 'Next Reminder Date',
  'Do Not Contact', 'Do Not Contact Reason'
];
const FOLLOWUP_REVIEW_FIELDS = [
  'Updated Support Details', 'Follow-Up Outcome', 'Eligibility Review Status',
  'Reviewed By', 'Review Date', 'PI Notes'
];
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
  'Parent Email', 'Parent Phone', 'Match Status', 'Match Confidence Score',
  'Match Reasons', 'Possible Consent Matches', 'Manual Consent ResponseID',
  'Manual Match Notes', 'Last Updated'
];
const UNMATCHED_CONSENT_HEADERS = [
  'Consent ResponseID', 'Submitted At', 'Child Full Name', 'Parent Full Name',
  'Parent Email', 'Parent Phone', 'Consent Status', 'Best Prescreening Match',
  'Best Match Score', 'Match Reasons', 'Review Status', 'Notes'
];

const EMAIL_TEMPLATES_SHEET = 'Email_Templates';
const EMAIL_OUTBOX_SHEET = 'Email_Outbox';
const EMAIL_LOG_SHEET = 'Email_Log';
const REPLY_LOG_SHEET = 'Reply_Log';
const REMINDER_QUEUE_SHEET = 'Reminder_Queue';
const POWER_AUTOMATE_CONFIG_SHEET = 'Power_Automate_Config';
const EMAIL_TEMPLATE_HEADERS = [
  'TemplateID', 'Template Name', 'Email Type', 'Subject Template', 'Body Template',
  'Active', 'Requires Approval', 'Approved By', 'Approved Date', 'Last Updated', 'Notes'
];
const EMAIL_OUTBOX_HEADERS = [
  'EmailJobID', 'EnrollmentID', 'PrescreeningID', 'ConsentID', 'Child Full Name',
  'Parent/Caretaker Name', 'Parent Email', 'Parent Phone', 'Cohort ID', 'Cohort Name',
  'Site', 'Program Term', 'Email Type', 'TemplateID', 'Email Subject', 'Email Body',
  'Email Body HTML', 'Approval Status', 'Approved By', 'Approved At', 'Send Status',
  'Scheduled Send At', 'Sent At', 'Power Automate Run ID', 'Power Automate Status',
  'Retry Count', 'Last Attempt At', 'Error Message', 'Created At', 'Updated At',
  'Do Not Send', 'Notes'
];
const EMAIL_LOG_HEADERS = [
  'LogID', 'Timestamp', 'EmailJobID', 'EnrollmentID', 'Parent Email', 'Child Full Name',
  'Email Type', 'Subject', 'Action', 'Status', 'Sent By / Mailbox',
  'Power Automate Run ID', 'Error Message', 'Raw Response'
];
const REPLY_LOG_HEADERS = [
  'ReplyID', 'EmailJobID', 'EnrollmentID', 'Parent Email', 'From Name', 'From Address',
  'Subject', 'Received At', 'Reply Preview', 'Reply Body', 'Has Attachments',
  'Power Automate Message ID', 'Follow Up Status Before Reply', 'Follow Up Status After Reply',
  'Reviewed By', 'Reviewed At', 'Notes'
];
const REMINDER_QUEUE_HEADERS = [
  'ReminderID', 'EmailJobID', 'EnrollmentID', 'Reminder Number', 'Reminder Type',
  'Scheduled Send At', 'Approval Status', 'Send Status', 'Sent At',
  'Power Automate Run ID', 'Error Message', 'Created At', 'Notes'
];
const POWER_AUTOMATE_CONFIG_HEADERS = ['Setting', 'Value', 'Notes'];
const EMAIL_FOLLOWUP_FIELDS = [
  'Email Workflow Status', 'Latest EmailJobID', 'Latest Email Type', 'Latest Email Subject',
  'Latest Email Approved At', 'Latest Email Sent At', 'Latest Reply Received At',
  'Reply Status', 'Reminder Status', 'Reminder Count', 'Next Reminder Date',
  'Do Not Contact', 'Do Not Contact Reason'
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
    .addItem('Initialize email workflow sheets', 'initializeEmailWorkflow')
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
  if (action === 'email_jobs') return jsonResponse_({ok: true, jobs: getApprovedEmailJobsForPowerAutomate(e && e.parameter && e.parameter.secret)});
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
    const action = String((e && e.parameter && e.parameter.action) || payload.action || '').toLowerCase();
    if (action === 'email_status') return jsonResponse_(receivePowerAutomateSendUpdate_(payload));
    if (action === 'email_reply') return jsonResponse_(receivePowerAutomateReply_(payload));
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
  ensureEmailWorkflowInfrastructure_(ss);

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
  // Intentionally do not rebuild workbook tabs from the web app.
  // QuestionPro/webhook processing and manual workbook refreshes update the Sheet database.
  // The dashboard only reads the current Sheet state so it cannot wipe or reset displayed data.
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  const professor = buildProfessorDashboardData_(false);
  const dashboardRows = readDashboardMetricRows_(ss);
  const matchReview = readObjectsOrEmpty_(ss, MATCH_REVIEW_SHEET);
  const unmatchedConsents = readObjectsOrEmpty_(ss, UNMATCHED_CONSENT_SHEET);
  const prescreens = readObjectsOrEmpty_(ss, SHEETS.PRESCREEN_CLEAN);
  const consents = readObjectsOrEmpty_(ss, SHEETS.CONSENT_CLEAN);
  const master = readObjectsOrEmpty_(ss, SHEETS.MASTER);
  const followups = readObjectsOrEmpty_(ss, SHEETS.FOLLOWUP);
  const ready = readObjectsOrEmpty_(ss, SHEETS.READY);
  const emailTemplates = readObjectsOrEmpty_(ss, EMAIL_TEMPLATES_SHEET);
  const emailOutbox = readObjectsOrEmpty_(ss, EMAIL_OUTBOX_SHEET);
  const emailLog = readObjectsOrEmpty_(ss, EMAIL_LOG_SHEET);
  const replyLog = readObjectsOrEmpty_(ss, REPLY_LOG_SHEET);
  const reminderQueue = readObjectsOrEmpty_(ss, REMINDER_QUEUE_SHEET);
  const payload = {
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
    cohorts: buildCohortSummaries_(ss, professor.participants, prescreens, consents),
    emailTemplates,
    emailOutbox,
    emailLog,
    replyLog,
    reminderQueue,
    emailMetrics: buildEmailMetrics_(emailOutbox, emailLog, replyLog, reminderQueue),
    rawCounts: {
      prescreening: readRawRows_(ss.getSheetByName(SHEETS.PRESCREEN_RAW)).length,
      consent: readRawRows_(ss.getSheetByName(SHEETS.CONSENT_RAW)).length
    },
    activity: buildActivityFeed_(prescreens, consents, followups, matchReview, unmatchedConsents, emailOutbox, replyLog),
    reports: buildReportData_(professor.participants, followups, matchReview, unmatchedConsents, ready, emailOutbox, replyLog, reminderQueue),
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

function buildActivityFeed_(prescreens, consents, followups, matchReview, unmatchedConsents, emailOutbox, replyLog) {
  const items = [];
  prescreens.slice(-8).forEach(row => items.push({type: 'Prescreening', title: row['Child Full Name'] || 'Prescreening submitted', detail: row['Parent Email'] || row['Parent/Caretaker Name'] || '', when: row['Submitted At'] || ''}));
  consents.slice(-8).forEach(row => items.push({type: 'Consent', title: row['Child Full Name'] || 'Consent submitted', detail: row['Parent Email'] || row['Parent Full Name'] || '', when: row['Submitted At'] || ''}));
  followups.filter(row => row['Follow Up Status'] && row['Follow Up Status'] !== 'Not Started').slice(-8).forEach(row => items.push({type: 'Follow-Up', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Follow Up Status'], when: row['Follow Up Completed Date'] || row['Email Sent Date'] || ''}));
  matchReview.slice(-8).forEach(row => items.push({type: 'Match Review', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Match Status'] || 'Needs Review', when: row['Last Updated'] || ''}));
  unmatchedConsents.slice(-8).forEach(row => items.push({type: 'Unmatched Consent', title: row['Child Full Name'] || row['Consent ResponseID'], detail: row['Best Prescreening Match'] || 'No confident match', when: row['Submitted At'] || ''}));
  (emailOutbox || []).slice(-8).forEach(row => items.push({type: 'Email', title: row['Child Full Name'] || row['EmailJobID'], detail: `${row['Email Type'] || 'Email'} • ${row['Send Status'] || row['Approval Status'] || ''}`, when: row['Updated At'] || row['Created At'] || ''}));
  (replyLog || []).slice(-8).forEach(row => items.push({type: 'Reply', title: row['Child Full Name'] || row['EnrollmentID'], detail: row['Reply Preview'] || row['Subject'] || '', when: row['Received At'] || ''}));
  return items.sort((a, b) => String(b.when || '').localeCompare(String(a.when || ''))).slice(0, 15);
}

function buildReportData_(participants, followups, matchReview, unmatchedConsents, ready, emailOutbox, replyLog, reminderQueue) {
  return {
    followupStatus: groupCounts_(followups, 'Follow Up Status'),
    consentStatus: groupCounts_(participants, 'consentStatus'),
    readiness: groupCounts_(participants, 'readyForDarts'),
    matchStatus: groupCounts_(participants, 'matchStatus'),
    reviewStatus: groupCounts_(participants, 'eligibilityReviewStatus'),
    emailSendStatus: groupCounts_(emailOutbox || [], 'Send Status'),
    emailApprovalStatus: groupCounts_(emailOutbox || [], 'Approval Status'),
    replyStatus: groupCounts_(replyLog || [], 'Follow Up Status After Reply'),
    reminderStatus: groupCounts_(reminderQueue || [], 'Send Status'),
    actionItems: [
      {label: 'Follow-ups not started', value: followups.filter(row => !row['Follow Up Status'] || row['Follow Up Status'] === 'Not Started').length},
      {label: 'Follow-ups awaiting response', value: followups.filter(row => row['Follow Up Status'] === 'Awaiting Response').length},
      {label: 'Match review needed', value: matchReview.length},
      {label: 'Unmatched consent records', value: unmatchedConsents.length},
      {label: 'Ready for DARTS export', value: ready.length},
      {label: 'Emails pending approval', value: (emailOutbox || []).filter(row => row['Approval Status'] === 'Pending Approval').length},
      {label: 'Email send failures', value: (emailOutbox || []).filter(row => row['Send Status'] === 'Failed').length},
      {label: 'Parent replies needing review', value: (replyLog || []).filter(row => !row['Reviewed At']).length}
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
      cohortId: row['Cohort ID'] || prescreen['Cohort ID'] || consent['Cohort ID'] || '',
      cohortName: row['Cohort Name'] || prescreen['Cohort Name'] || consent['Cohort Name'] || '',
      site: row['Site'] || prescreen['Site'] || consent['Site'] || '',
      programTerm: row['Program Term'] || prescreen['Program Term'] || consent['Program Term'] || '',
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
    'Cohort ID': prescreen['Cohort ID'] || '',
    'Cohort Name': prescreen['Cohort Name'] || '',
    'Site': prescreen['Site'] || '',
    'Program Term': prescreen['Program Term'] || '',
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
  const allowed = ['Follow Up Status', 'Email Drafted', 'Email Sent Date', 'Follow Up Completed Date', 'Assigned To', 'Notes', ...FOLLOWUP_REVIEW_FIELDS, ...EMAIL_FOLLOWUP_FIELDS];
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
  FOLLOWUP_REVIEW_FIELDS.concat(EMAIL_FOLLOWUP_FIELDS).forEach(field => {
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
    ['Needs Review', countWhere_(ss, SHEETS.MASTER, 'Ready for DARTS', 'Review')],
    ['Emails Pending Approval', countWhere_(ss, EMAIL_OUTBOX_SHEET, 'Approval Status', 'Pending Approval')],
    ['Emails Ready To Send', countWhere_(ss, EMAIL_OUTBOX_SHEET, 'Send Status', 'Ready To Send')],
    ['Emails Sent', countWhere_(ss, EMAIL_OUTBOX_SHEET, 'Send Status', 'Sent')],
    ['Replies Received', countRows_(ss, REPLY_LOG_SHEET)]
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
      'Cohort ID': p['Cohort ID'] || (consent && consent['Cohort ID']) || '',
      'Cohort Name': p['Cohort Name'] || (consent && consent['Cohort Name']) || '',
      'Site': p['Site'] || (consent && consent['Site']) || '',
      'Program Term': p['Program Term'] || (consent && consent['Program Term']) || '',
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
      'PI Notes': manual['PI Notes'] || '',
      'Email Workflow Status': manual['Email Workflow Status'] || '',
      'Latest EmailJobID': manual['Latest EmailJobID'] || '',
      'Latest Email Type': manual['Latest Email Type'] || '',
      'Latest Email Subject': manual['Latest Email Subject'] || '',
      'Latest Email Approved At': manual['Latest Email Approved At'] || '',
      'Latest Email Sent At': manual['Latest Email Sent At'] || '',
      'Latest Reply Received At': manual['Latest Reply Received At'] || '',
      'Reply Status': manual['Reply Status'] || 'No Reply',
      'Reminder Status': manual['Reminder Status'] || '',
      'Reminder Count': manual['Reminder Count'] || 0,
      'Next Reminder Date': manual['Next Reminder Date'] || '',
      'Do Not Contact': manual['Do Not Contact'] || 'No',
      'Do Not Contact Reason': manual['Do Not Contact Reason'] || ''
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


// ---------- Email workflow / Power Automate handoff ----------

function initializeEmailWorkflow() {
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  return buildAppDashboardData_(false);
}

function ensureEmailWorkflowInfrastructure_(ss) {
  const templates = ensureSheetWithHeaders_(ss, EMAIL_TEMPLATES_SHEET, EMAIL_TEMPLATE_HEADERS);
  ensureSheetWithHeaders_(ss, EMAIL_OUTBOX_SHEET, EMAIL_OUTBOX_HEADERS);
  ensureSheetWithHeaders_(ss, EMAIL_LOG_SHEET, EMAIL_LOG_HEADERS);
  ensureSheetWithHeaders_(ss, REPLY_LOG_SHEET, REPLY_LOG_HEADERS);
  ensureSheetWithHeaders_(ss, REMINDER_QUEUE_SHEET, REMINDER_QUEUE_HEADERS);
  ensureSheetWithHeaders_(ss, POWER_AUTOMATE_CONFIG_SHEET, POWER_AUTOMATE_CONFIG_HEADERS);
  appendMissingHeaders_(ss.getSheetByName(SHEETS.FOLLOWUP), EMAIL_FOLLOWUP_FIELDS);
  seedDefaultEmailTemplates_(templates);
}

function seedDefaultEmailTemplates_(sheet) {
  const rows = readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW);
  if (rows.some(row => row['TemplateID'] === 'TPL-FOLLOWUP-INITIAL')) return;
  sheet.appendRow([
    'TPL-FOLLOWUP-INITIAL',
    'Initial no-response clarification follow-up',
    'Initial Follow-Up',
    'Gaming4Good follow-up for {{child_first_name}}',
    [
      'Dear {{parent_first_name}},',
      '',
      'Thank you for completing the Gaming4Good prescreening form for {{child_first_name}}.',
      '',
      'We are following up because the prescreening response indicated that {{child_first_name}} may not currently identify as neurodivergent or have a formal disability/developmental condition/learning difference diagnosis. We would like to better understand whether there are any support needs, learning differences, school accommodations, services, or other information that may help the research team determine eligibility.',
      '',
      'If you are comfortable sharing, please reply with any additional details about current supports, diagnoses, accommodations, or learning needs that may not have been fully captured in the form.',
      '',
      'Thank you,',
      'The Gaming4Good Research Team',
      '',
      'Reference: {{email_job_id}}'
    ].join('\n'),
    'Yes', 'Yes', '', '', new Date(), 'Default template created by Apps Script. Review and approve before production use.'
  ]);
}

function getEmailTemplates_(ss) {
  ensureEmailWorkflowInfrastructure_(ss);
  return readObjects_(ss.getSheetByName(EMAIL_TEMPLATES_SHEET), CLEAN_HEADER_ROW, DATA_START_ROW)
    .filter(row => String(row['Active'] || '').toLowerCase() !== 'no');
}

function previewFollowupEmail(request) {
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  const template = findEmailTemplate_(ss, request && request.templateId);
  const context = buildEmailMergeContext_(ss, request && request.enrollmentId, request && request.emailJobId);
  return buildEmailDraftPayload_(template, context);
}

function saveEmailDraft(request) {
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  const enrollmentId = String(request && request.enrollmentId || '').trim();
  if (!enrollmentId) throw new Error('EnrollmentID is required to create an email draft.');
  const template = findEmailTemplate_(ss, request && request.templateId);
  const existingJobId = String(request && request.emailJobId || '').trim();
  const emailJobId = existingJobId || nextId_('EMAIL');
  const context = buildEmailMergeContext_(ss, enrollmentId, emailJobId);
  const draft = buildEmailDraftPayload_(template, context);
  const now = new Date();
  const record = {
    'EmailJobID': emailJobId,
    'EnrollmentID': enrollmentId,
    'PrescreeningID': context.prescreeningId,
    'ConsentID': context.consentId,
    'Child Full Name': context.childFullName,
    'Parent/Caretaker Name': context.parentFullName,
    'Parent Email': context.parentEmail,
    'Parent Phone': context.parentPhone,
    'Cohort ID': context.cohortId,
    'Cohort Name': context.cohortName,
    'Site': context.site,
    'Program Term': context.programTerm,
    'Email Type': request && request.emailType || template['Email Type'] || 'Initial Follow-Up',
    'TemplateID': template['TemplateID'],
    'Email Subject': request && request.subject || draft.subject,
    'Email Body': request && request.body || draft.body,
    'Email Body HTML': request && request.bodyHtml || draft.bodyHtml,
    'Approval Status': 'Pending Approval',
    'Approved By': '',
    'Approved At': '',
    'Send Status': 'Draft',
    'Scheduled Send At': request && request.scheduledSendAt || '',
    'Sent At': '',
    'Power Automate Run ID': '',
    'Power Automate Status': '',
    'Retry Count': 0,
    'Last Attempt At': '',
    'Error Message': '',
    'Created At': existingJobId ? '' : now,
    'Updated At': now,
    'Do Not Send': 'No',
    'Notes': request && request.notes || ''
  };
  upsertByKey_(ss.getSheetByName(EMAIL_OUTBOX_SHEET), EMAIL_OUTBOX_HEADERS, 'EmailJobID', emailJobId, record);
  appendEmailLog_(ss, record, 'Draft Created', 'Draft', '', '');
  updateFollowupEmailFields_(ss, enrollmentId, {
    'Email Workflow Status': 'Draft Created',
    'Email Drafted': 'Yes',
    'Latest EmailJobID': emailJobId,
    'Latest Email Type': record['Email Type'],
    'Latest Email Subject': record['Email Subject']
  });
  return buildAppDashboardData_(false);
}

function approveEmailForSending(request) {
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  const emailJobId = String(request && request.emailJobId || '').trim();
  if (!emailJobId) throw new Error('EmailJobID is required for approval.');
  const sheet = ss.getSheetByName(EMAIL_OUTBOX_SHEET);
  const rowInfo = findRowByKey_(sheet, 'EmailJobID', emailJobId);
  if (!rowInfo.rowNumber) throw new Error(`Email job not found: ${emailJobId}`);
  const record = rowInfo.object;
  if (record['Do Not Send'] === 'Yes') throw new Error('This email job is marked Do Not Send.');
  if (!record['Parent Email']) throw new Error('Parent Email is required before approval.');
  const scheduled = String(request && request.scheduledSendAt || record['Scheduled Send At'] || '').trim();
  const now = new Date();
  const updates = {
    'Approval Status': 'Approved',
    'Approved By': request && request.approvedBy || Session.getActiveUser().getEmail() || 'Dashboard Admin',
    'Approved At': now,
    'Send Status': scheduled ? 'Scheduled' : 'Ready To Send',
    'Scheduled Send At': scheduled,
    'Updated At': now,
    'Notes': request && request.notes !== undefined ? request.notes : record['Notes']
  };
  updateRowByHeader_(sheet, rowInfo.rowNumber, updates);
  appendEmailLog_(ss, Object.assign({}, record, updates), 'Approved', updates['Send Status'], '', '');
  updateFollowupEmailFields_(ss, record['EnrollmentID'], {
    'Email Workflow Status': updates['Send Status'],
    'Latest Email Approved At': now,
    'Latest EmailJobID': emailJobId,
    'Latest Email Type': record['Email Type'],
    'Latest Email Subject': record['Email Subject']
  });
  return buildAppDashboardData_(false);
}

function cancelEmailJob(request) {
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  const emailJobId = String(request && request.emailJobId || '').trim();
  const sheet = ss.getSheetByName(EMAIL_OUTBOX_SHEET);
  const rowInfo = findRowByKey_(sheet, 'EmailJobID', emailJobId);
  if (!rowInfo.rowNumber) throw new Error(`Email job not found: ${emailJobId}`);
  const updates = {'Send Status': 'Cancelled', 'Do Not Send': 'Yes', 'Updated At': new Date(), 'Notes': request && request.reason || rowInfo.object['Notes'] || ''};
  updateRowByHeader_(sheet, rowInfo.rowNumber, updates);
  appendEmailLog_(ss, Object.assign({}, rowInfo.object, updates), 'Cancelled', 'Cancelled', '', '');
  updateFollowupEmailFields_(ss, rowInfo.object['EnrollmentID'], {'Email Workflow Status': 'Cancelled'});
  return buildAppDashboardData_(false);
}

function markDoNotContact(request) {
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  const enrollmentId = String(request && request.enrollmentId || '').trim();
  if (!enrollmentId) throw new Error('EnrollmentID is required.');
  updateFollowupEmailFields_(ss, enrollmentId, {'Do Not Contact': 'Yes', 'Do Not Contact Reason': request && request.reason || '', 'Email Workflow Status': 'Do Not Contact'});
  return buildAppDashboardData_(false);
}

function getApprovedEmailJobsForPowerAutomate(secret) {
  const ss = getSpreadsheet_();
  validatePowerAutomateSecret_(secret);
  ensureEmailWorkflowInfrastructure_(ss);
  const now = new Date();
  return readObjects_(ss.getSheetByName(EMAIL_OUTBOX_SHEET), CLEAN_HEADER_ROW, DATA_START_ROW)
    .filter(row => row['Approval Status'] === 'Approved')
    .filter(row => ['Ready To Send', 'Scheduled', 'Failed'].includes(String(row['Send Status'] || '')))
    .filter(row => row['Do Not Send'] !== 'Yes' && row['Sent At'] === '')
    .filter(row => !row['Scheduled Send At'] || new Date(row['Scheduled Send At']) <= now);
}

function receivePowerAutomateSendUpdate_(payload) {
  const ss = getSpreadsheet_();
  validatePowerAutomateSecret_(payload && (payload.secret || payload.callbackSecret));
  ensureEmailWorkflowInfrastructure_(ss);
  const emailJobId = String(payload.emailJobId || payload.EmailJobID || '').trim();
  if (!emailJobId) throw new Error('Power Automate callback missing emailJobId.');
  const sheet = ss.getSheetByName(EMAIL_OUTBOX_SHEET);
  const rowInfo = findRowByKey_(sheet, 'EmailJobID', emailJobId);
  if (!rowInfo.rowNumber) throw new Error(`Email job not found: ${emailJobId}`);
  const status = String(payload.status || payload.sendStatus || 'Sent');
  const sent = /sent|success/i.test(status);
  const updates = {
    'Send Status': sent ? 'Sent' : 'Failed',
    'Sent At': sent ? (payload.sentAt || new Date()) : rowInfo.object['Sent At'],
    'Power Automate Run ID': payload.runId || payload.powerAutomateRunId || rowInfo.object['Power Automate Run ID'] || '',
    'Power Automate Status': status,
    'Retry Count': Number(rowInfo.object['Retry Count'] || 0) + (sent ? 0 : 1),
    'Last Attempt At': payload.attemptedAt || new Date(),
    'Error Message': payload.error || payload.errorMessage || '',
    'Updated At': new Date()
  };
  updateRowByHeader_(sheet, rowInfo.rowNumber, updates);
  appendEmailLog_(ss, Object.assign({}, rowInfo.object, updates), sent ? 'Sent' : 'Failed', updates['Send Status'], updates['Power Automate Run ID'], JSON.stringify(payload));
  updateFollowupEmailFields_(ss, rowInfo.object['EnrollmentID'], {
    'Email Workflow Status': sent ? 'Awaiting Response' : 'Send Failed',
    'Follow Up Status': sent ? 'Awaiting Response' : rowInfo.object['Follow Up Status'],
    'Email Sent Date': sent ? updates['Sent At'] : '',
    'Latest Email Sent At': sent ? updates['Sent At'] : '',
    'Latest EmailJobID': emailJobId
  });
  return {ok: true, emailJobId, status: updates['Send Status']};
}

function receivePowerAutomateReply_(payload) {
  const ss = getSpreadsheet_();
  validatePowerAutomateSecret_(payload && (payload.secret || payload.callbackSecret));
  ensureEmailWorkflowInfrastructure_(ss);
  const emailJobId = String(payload.emailJobId || payload.EmailJobID || '').trim();
  const enrollmentId = String(payload.enrollmentId || payload.EnrollmentID || '').trim() || enrollmentIdForEmailJob_(ss, emailJobId);
  const followupBefore = getFollowupStatus_(ss, enrollmentId);
  const reply = {
    'ReplyID': payload.replyId || nextId_('REPLY'),
    'EmailJobID': emailJobId,
    'EnrollmentID': enrollmentId,
    'Parent Email': payload.parentEmail || payload.fromAddress || payload.from || '',
    'From Name': payload.fromName || '',
    'From Address': payload.fromAddress || payload.from || '',
    'Subject': payload.subject || '',
    'Received At': payload.receivedAt || new Date(),
    'Reply Preview': payload.bodyPreview || String(payload.bodyText || payload.body || '').slice(0, 300),
    'Reply Body': payload.bodyText || payload.body || '',
    'Has Attachments': payload.hasAttachments || '',
    'Power Automate Message ID': payload.messageId || '',
    'Follow Up Status Before Reply': followupBefore,
    'Follow Up Status After Reply': 'Parent Replied',
    'Reviewed By': '',
    'Reviewed At': '',
    'Notes': payload.notes || ''
  };
  appendObjectByHeaders_(ss.getSheetByName(REPLY_LOG_SHEET), REPLY_LOG_HEADERS, reply);
  appendEmailLog_(ss, {'EmailJobID': emailJobId, 'EnrollmentID': enrollmentId, 'Parent Email': reply['Parent Email'], 'Child Full Name': '', 'Email Type': 'Reply', 'Email Subject': reply['Subject']}, 'Reply Received', 'Parent Replied', payload.runId || '', JSON.stringify(payload));
  updateFollowupEmailFields_(ss, enrollmentId, {'Follow Up Status': 'Parent Replied', 'Email Workflow Status': 'Parent Replied', 'Latest Reply Received At': reply['Received At'], 'Reply Status': 'Reply Received'});
  return {ok: true, emailJobId, enrollmentId, status: 'Parent Replied'};
}

function scheduleReminder(request) {
  const ss = getSpreadsheet_();
  ensureEmailWorkflowInfrastructure_(ss);
  const reminderId = nextId_('REMINDER');
  const record = {
    'ReminderID': reminderId,
    'EmailJobID': request.emailJobId || '',
    'EnrollmentID': request.enrollmentId || '',
    'Reminder Number': request.reminderNumber || 1,
    'Reminder Type': request.reminderType || 'Follow-Up Reminder',
    'Scheduled Send At': request.scheduledSendAt || '',
    'Approval Status': 'Pending Approval',
    'Send Status': 'Draft',
    'Sent At': '',
    'Power Automate Run ID': '',
    'Error Message': '',
    'Created At': new Date(),
    'Notes': request.notes || ''
  };
  appendObjectByHeaders_(ss.getSheetByName(REMINDER_QUEUE_SHEET), REMINDER_QUEUE_HEADERS, record);
  updateFollowupEmailFields_(ss, record['EnrollmentID'], {'Reminder Status': 'Reminder Drafted', 'Next Reminder Date': record['Scheduled Send At']});
  return buildAppDashboardData_(false);
}

function buildEmailMetrics_(outbox, logs, replies, reminders) {
  outbox = outbox || [];
  replies = replies || [];
  reminders = reminders || [];
  return {
    drafts: outbox.filter(row => row['Send Status'] === 'Draft').length,
    pendingApproval: outbox.filter(row => row['Approval Status'] === 'Pending Approval').length,
    readyToSend: outbox.filter(row => row['Send Status'] === 'Ready To Send').length,
    scheduled: outbox.filter(row => row['Send Status'] === 'Scheduled').length,
    sent: outbox.filter(row => row['Send Status'] === 'Sent').length,
    failed: outbox.filter(row => row['Send Status'] === 'Failed').length,
    repliesReceived: replies.length,
    remindersPending: reminders.filter(row => row['Send Status'] !== 'Sent').length
  };
}

function findEmailTemplate_(ss, templateId) {
  const templates = getEmailTemplates_(ss);
  return templates.find(row => row['TemplateID'] === templateId) || templates[0] || {};
}

function buildEmailMergeContext_(ss, enrollmentId, emailJobId) {
  const participants = buildProfessorDashboardData_(false).participants || [];
  const participant = participants.find(p => p.enrollmentId === enrollmentId);
  if (!participant) throw new Error(`EnrollmentID not found: ${enrollmentId}`);
  const nameParts = String(participant.parentName || '').trim().split(/\s+/).filter(Boolean);
  const childParts = String(participant.childName || '').trim().split(/\s+/).filter(Boolean);
  return {
    emailJobId: emailJobId || nextId_('EMAIL-PREVIEW'),
    enrollmentId: participant.enrollmentId,
    prescreeningId: '',
    consentId: '',
    childFullName: participant.childName || '',
    childFirstName: childParts[0] || participant.childName || '',
    parentFullName: participant.parentName || '',
    parentFirstName: nameParts[0] || participant.parentName || 'there',
    parentEmail: participant.parentEmail || '',
    parentPhone: participant.parentPhone || '',
    cohortId: participant.cohortId || '',
    cohortName: participant.cohortName || '',
    site: participant.site || '',
    programTerm: participant.programTerm || '',
    followupReason: participant.followUpStatus || '',
    supportDetails: participant.supportDetails || ''
  };
}

function buildEmailDraftPayload_(template, context) {
  const subject = mergeTemplate_(template['Subject Template'] || 'Gaming4Good follow-up for {{child_first_name}}', context);
  const body = mergeTemplate_(template['Body Template'] || '', context);
  return {templateId: template['TemplateID'] || '', emailType: template['Email Type'] || '', subject, body, bodyHtml: textToHtml_(body), context};
}

function mergeTemplate_(template, context) {
  const fields = {
    email_job_id: context.emailJobId,
    enrollment_id: context.enrollmentId,
    parent_first_name: context.parentFirstName,
    parent_full_name: context.parentFullName,
    parent_email: context.parentEmail,
    child_first_name: context.childFirstName,
    child_full_name: context.childFullName,
    cohort_id: context.cohortId,
    cohort_name: context.cohortName,
    site: context.site,
    program_term: context.programTerm,
    followup_reason: context.followupReason,
    support_details: context.supportDetails
  };
  return String(template || '').replace(/{{\s*([a-zA-Z0-9_]+)\s*}}/g, (_, key) => fields[String(key).toLowerCase()] || '');
}

function textToHtml_(text) {
  return String(text || '').split(/\n{2,}/).map(p => `<p>${String(p).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>')}</p>`).join('');
}

function updateFollowupEmailFields_(ss, enrollmentId, updates) {
  const sheet = ss.getSheetByName(SHEETS.FOLLOWUP);
  if (!sheet || !enrollmentId) return;
  appendMissingHeaders_(sheet, Object.keys(updates));
  const rowInfo = findRowByKey_(sheet, 'EnrollmentID', enrollmentId);
  if (rowInfo.rowNumber) updateRowByHeader_(sheet, rowInfo.rowNumber, updates);
  if (updates['Follow Up Status'] !== undefined) {
    const masterSheet = ss.getSheetByName(SHEETS.MASTER);
    const masterRow = findRowByKey_(masterSheet, 'EnrollmentID', enrollmentId);
    if (masterRow.rowNumber) updateRowByHeader_(masterSheet, masterRow.rowNumber, {'Follow Up Status': updates['Follow Up Status']});
  }
}

function getFollowupStatus_(ss, enrollmentId) {
  const rowInfo = findRowByKey_(ss.getSheetByName(SHEETS.FOLLOWUP), 'EnrollmentID', enrollmentId);
  return rowInfo.object && rowInfo.object['Follow Up Status'] || '';
}

function enrollmentIdForEmailJob_(ss, emailJobId) {
  const rowInfo = findRowByKey_(ss.getSheetByName(EMAIL_OUTBOX_SHEET), 'EmailJobID', emailJobId);
  return rowInfo.object && rowInfo.object['EnrollmentID'] || '';
}

function appendEmailLog_(ss, emailRecord, action, status, runId, rawResponse) {
  appendObjectByHeaders_(ss.getSheetByName(EMAIL_LOG_SHEET), EMAIL_LOG_HEADERS, {
    'LogID': nextId_('LOG'),
    'Timestamp': new Date(),
    'EmailJobID': emailRecord['EmailJobID'] || '',
    'EnrollmentID': emailRecord['EnrollmentID'] || '',
    'Parent Email': emailRecord['Parent Email'] || '',
    'Child Full Name': emailRecord['Child Full Name'] || '',
    'Email Type': emailRecord['Email Type'] || '',
    'Subject': emailRecord['Email Subject'] || emailRecord['Subject'] || '',
    'Action': action,
    'Status': status,
    'Sent By / Mailbox': '',
    'Power Automate Run ID': runId || '',
    'Error Message': emailRecord['Error Message'] || '',
    'Raw Response': rawResponse || ''
  });
}

function validatePowerAutomateSecret_(provided) {
  const expected = PropertiesService.getScriptProperties().getProperty('POWER_AUTOMATE_SECRET') || getPowerAutomateConfigValue_('PowerAutomateCallbackSecret');
  if (!expected) return true;
  if (String(provided || '') !== String(expected)) throw new Error('Invalid Power Automate callback secret.');
  return true;
}

function getPowerAutomateConfigValue_(setting) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(POWER_AUTOMATE_CONFIG_SHEET);
  if (!sheet) return '';
  const row = readObjects_(sheet, CLEAN_HEADER_ROW, DATA_START_ROW).find(r => r['Setting'] === setting);
  return row ? row['Value'] : '';
}

function nextId_(prefix) {
  return `${prefix}-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss')}-${Math.floor(Math.random() * 10000)}`;
}

function findRowByKey_(sheet, keyField, keyValue) {
  if (!sheet || !keyValue) return {rowNumber: 0, object: {}};
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const keyCol = headers.indexOf(keyField);
  if (keyCol < 0) return {rowNumber: 0, object: {}};
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyCol]) === String(keyValue)) {
      const object = {};
      headers.forEach((h, idx) => { if (h) object[h] = values[i][idx]; });
      return {rowNumber: i + 1, object};
    }
  }
  return {rowNumber: 0, object: {}};
}

function updateRowByHeader_(sheet, rowNumber, updates) {
  appendMissingHeaders_(sheet, Object.keys(updates));
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  Object.entries(updates).forEach(([field, value]) => {
    const col = headers.indexOf(field) + 1;
    if (col > 0) sheet.getRange(rowNumber, col).setValue(value);
  });
}

function upsertByKey_(sheet, headers, keyField, keyValue, record) {
  appendMissingHeaders_(sheet, headers);
  const rowInfo = findRowByKey_(sheet, keyField, keyValue);
  if (rowInfo.rowNumber) {
    if (!record['Created At']) delete record['Created At'];
    updateRowByHeader_(sheet, rowInfo.rowNumber, record);
  } else {
    appendObjectByHeaders_(sheet, headers, record);
  }
}

function appendObjectByHeaders_(sheet, headers, record) {
  appendMissingHeaders_(sheet, headers);
  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  sheet.appendRow(currentHeaders.map(header => record[header] === undefined ? '' : record[header]));
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
  [SHEETS.PRESCREEN_CLEAN, SHEETS.CONSENT_CLEAN, SHEETS.MASTER, SHEETS.FOLLOWUP, SHEETS.READY, MATCH_REVIEW_SHEET, UNMATCHED_CONSENT_SHEET]
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
  if (![SHEETS.PRESCREEN_CLEAN, SHEETS.FOLLOWUP].includes(sheet.getName())) return;
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
