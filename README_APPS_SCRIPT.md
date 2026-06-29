# Gaming4Good Enrollment Tracker Apps Script

This repository contains the Phase 1 Google Apps Script automation for the existing Google Sheet tracker. In the Apps Script editor, create/paste only the files `Code.gs` and `Sidebar.html`.

## How the workbook is intended to work

- `Prescreening_Raw` and `Consent_Raw` are the audit trail. Do not edit these tabs manually after import/webhook receipt.
- `Prescreening_Clean` and `Consent_Clean` are rebuilt by the script from raw QuestionPro data.
- `Master_Enrollment` keeps one row per enrollment record, matching consent responses back to prescreening records by child name plus parent email/phone where possible.
- `Followup_Queue` is the working list for the research team. The script creates rows only when follow-up appears needed; staff update follow-up status/date/assignee/notes there.
- `Ready_For_DARTS` contains participants whose prescreening, required follow-up, and consent are complete.
- `Dashboard` is refreshed with high-level counts.
- `Config` stores defaults such as the default follow-up, consent, and enrollment statuses.

## Phase 1 operating workflow

1. Import QuestionPro exports into the raw tabs or deploy `doPost` as a web app for webhook intake.
2. In Google Sheets, use **G4G Tracker → Open tracker sidebar** or **G4G Tracker → Refresh tracker from raw tabs**.
3. Review `Followup_Queue` for rows with `Follow Up Status = Not Started`.
4. Manually contact parents outside this app. Phase 1 does not send email.
5. Update `Followup_Queue` fields such as `Email Drafted`, `Email Sent Date`, `Follow Up Status`, `Follow Up Completed Date`, `Assigned To`, and `Notes`.
6. Refresh the tracker again after consent imports or follow-up status changes.
7. Use `Ready_For_DARTS` when `Ready for DARTS = Yes` records are needed for downstream enrollment/export.

## Testing

Use **G4G Tracker → Run self-test with sample rows** to append a sample prescreening raw row, refresh the tracker, and verify that the child appears in `Followup_Queue`.

Suggested manual checks:

- A prescreening row with a neurodivergent/disability/support response of `Yes` appears in `Followup_Queue`.
- Updating that queue row to `Completed` is preserved after refresh.
- A matching consent row changes `Consent Status` in `Master_Enrollment` to `Completed`.
- A participant is listed in `Ready_For_DARTS` only when consent is complete and required follow-up is either complete or not needed.

## Apps Script file layout

This project intentionally includes only `.gs` and `.html` source files for the Apps Script editor. Configure deployment settings directly in the Apps Script UI when you publish the web app.
