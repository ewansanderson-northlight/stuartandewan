/* Paste into Apps Script, then deploy as a Web app.
   Execute as: Me. Who has access: Anyone. */
const SPREADSHEET_ID = '14MrafqjgL_m9RR5uGzYnLFT3TuWAEcPFwchpbLUFURQ';
const SHEET_NAME = 'Guest interest';
const HEADERS = ['Submitted at', 'Name', 'Attendees', 'Plus-one name', 'Email', 'Postal address', 'WhatsApp number', 'Children attending?', 'Likely to attend?'];

function doGet() { return ContentService.createTextOutput('Wedding guest-interest endpoint is ready.'); }

function doPost(e) {
  const sheet = getOrCreateSheet_();
  const fields = e.parameter;
  sheet.appendRow([fields.submittedAt || new Date().toISOString(), fields.name || '', fields.attendees || '', fields.plusOneName || '', fields.email || '', fields.address || '', fields.phone || '', fields.childrenAttending || '', fields.attendance || '']);
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.setFrozenRows(1); sheet.autoResizeColumns(1, HEADERS.length);
  }
  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (currentHeaders[2] !== 'Attendees') sheet.insertColumnsBefore(3, 2);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#315e61').setFontColor('#ffffff');
  return sheet;
}

function migrateGuestInterestSheet() { getOrCreateSheet_(); }
