/**
 * Serves the HTML file for the Web App.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Global Refunds Dashboard 2025')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Fetches data from the Google Sheet "Tracker" tab.
 */
function getSheetData() {
  const SHEET_NAME = 'Tracker';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" not found. Please check the tab name.`);
  }

  const data = sheet.getDataRange().getDisplayValues();
  return {
    headers: data[0],
    rows: data.slice(1)
  };
}

/**
 * NEW: Fetches the reservations data for the Heatmap Table.
 */
function getReservationsData() {
  const SHEET_NAME = 'No.of Resas/month/country';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  }

  const data = sheet.getDataRange().getDisplayValues();
  return {
    headers: data[0],
    rows: data.slice(1)
  };
}