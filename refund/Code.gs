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

  const range = sheet.getDataRange();
  const data = range.getDisplayValues();
  const richText = range.getRichTextValues();
  const formulas = range.getFormulas();

  // Find the column index for "Ticket ID"
  const ticketIdColIdx = data[0].findIndex(h => h.toLowerCase().includes('ticket id') || h.toLowerCase() === 'ticket');

  const ticketLinks = [];
  if (ticketIdColIdx !== -1) {
    for (let i = 1; i < data.length; i++) {
      // Try to get standard hyperlink, fallback to parsing =HYPERLINK() formula
      let link = richText[i][ticketIdColIdx].getLinkUrl();
      if (!link) {
        let formula = formulas[i][ticketIdColIdx];
        if (formula && formula.toUpperCase().includes('HYPERLINK')) {
          let match = formula.match(/HYPERLINK\(\s*"([^"]+)"/i);
          if (match) link = match[1];
        }
      }
      ticketLinks.push(link || null);
    }
  }

  return {
    headers: data[0],
    rows: data.slice(1),
    ticketLinks: ticketLinks
  };
}

/**
 * Fetches the reservations data for the Heatmap Table.
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