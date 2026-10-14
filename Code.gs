const SPREADSHEET_ID = '1IoO4DIkQYbmrWh2BH9B9ItAyrivcK1KJ7OWya_UuEv4';
const SHEET_NAME = 'AccessCodes';

function doGet(e) {
  const page = e && e.parameter && e.parameter.page === 'dashboard' ? 'Dashboard' : 'Index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'Dashboard' ? 'Rohansweb Dashboard' : 'Rohansweb Special Access')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once. Creates the AccessCodes tab and headings if missing. */
function setupSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['Code', 'Name', 'Active']]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 3);
  }
  return 'Ready: ' + ss.getUrl();
}

/** Validates one access code. Only the matching display name is returned. */
function checkAccessCode(code) {
  if (typeof code !== 'string' || !code.trim() || code.length > 200) {
    return { ok: false, message: 'Enter your access code.' };
  }
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { ok: false, message: 'No access codes have been configured.' };
  }
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues();
  for (const row of rows) {
    const savedCode = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const active = String(row[2] || '').trim().toLowerCase();
    if (savedCode && savedCode === code.trim() && name &&
        ['true', 'yes', 'active', '1'].includes(active)) {
      return { ok: true, name: name };
    }
  }
  return { ok: false, message: 'That access code was not recognised.' };
}
