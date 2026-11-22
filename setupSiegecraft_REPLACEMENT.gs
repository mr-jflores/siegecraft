function setupSiegecraft() {
  // Works from BOTH standalone Apps Script projects and Sheets-bound projects.
  // Reuses the same spreadsheet on later runs, including already-saved students.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const existingId = props.getProperty('SIEGECRAFT_SHEET_ID');
    let book;
    let created = false;

    if (existingId) {
      // Protect existing beta progress: NEVER create a second database simply
      // because a previously configured spreadsheet cannot be opened.
      try {
        book = SpreadsheetApp.openById(existingId);
      } catch (err) {
        throw new Error('Siegecraft is already linked to a spreadsheet, but it could not be opened. ' +
          'Check your access to this spreadsheet: https://docs.google.com/spreadsheets/d/' + existingId + '/edit');
      }
    } else {
      book = SpreadsheetApp.getActiveSpreadsheet();
      if (!book) {
        book = SpreadsheetApp.create('Siegecraft – Student Progress');
        created = true;
      }
    }

    let sheet = book.getSheetByName(SHEET_NAME);
    if (!sheet) {
      // New spreadsheets have a blank 'Sheet1' tab; rename rather than leave
      // an unused extra tab. Existing sheets are never renamed or removed.
      const tabs = book.getSheets();
      if (created && tabs.length === 1 && tabs[0].getLastRow() === 0) {
        sheet = tabs[0];
        sheet.setName(SHEET_NAME);
      } else {
        sheet = book.insertSheet(SHEET_NAME);
      }
    }

    if (sheet.getLastRow() === 0) sheet.appendRow(COLUMNS);
    else sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.autoResizeColumns(1, 10);

    // Store only after successful setup, then return the URL in the log.
    props.setProperty('SIEGECRAFT_SHEET_ID', book.getId());
    const message = 'Siegecraft spreadsheet ready: ' + book.getUrl() +
      ' | Student data tab: ' + SHEET_NAME;
    Logger.log(message);
    return message;
  } finally {
    lock.releaseLock();
  }
}
