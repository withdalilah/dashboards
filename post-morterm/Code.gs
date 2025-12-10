// @ts-nocheck
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const column = range.getColumn();
  const row = range.getRow();

  // Only act on Column A (1)
  if (column === 1) {
    const value = range.getValue();

    // Check if it's a URL and not already a formula
    if (typeof value === "string" && value.startsWith("http")) {
      const displayText = "Case " + (row - 1); // Adjust for header row
      const formula = `=HYPERLINK("${value}", "${displayText}")`;
      sheet.getRange(row, column).setFormula(formula);
    }
  }
}

function onFormSubmit(e) {
  const responses = e.namedValues;
  const timestamp = responses["Timestamp"][0];
  // const email = responses["Email Address"][0]; // Adjust if your form field name is different
  const caseId = responses["Case ID"] ? responses["Case ID"][0] : "No ID Case"; // Adjust as needed
  const caseIdColumnIndex = 0;
  const folderId = "1_gQPZoMSh7t9JHEQwFbURA48GxYL86ew"; // <-- Replace with your Google Drive folder ID

  // 1. Create a Google Doc from form answers
  const doc = DocumentApp.create(`Response - ${caseId}`);
  const body = doc.getBody();
  body.appendParagraph("📄 Case Summary").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`🕓 Submitted: ${timestamp}`);
  body.appendParagraph(`📧 Respondent: ENTAH LAH\n`);

  for (const question in responses) {
    if (question !== "Timestamp") {
      const answer = responses[question][0];
      body.appendParagraph(`❓ ${question}`).setBold(true);
      body.appendParagraph(`💬 ${answer}\n`);
    }
  }

  doc.saveAndClose();

  // 2. Convert to PDF
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF);
  const pdfFile = DriveApp.getFolderById(folderId).createFile(pdfBlob);
  pdfFile.setName(`Response - ${caseId}.pdf`);

  // 3. Set PDF permissions (optional)
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 4. Log PDF link to sheet + replace form link
  const ss = SpreadsheetApp.openById("175WBoQpRxSaOBOQUSugM9OGA0JJfKYRTIkhUI6RiErw"); // <-- Put your actual spreadsheet ID here
  const sheet = ss.getSheetByName("Feedbacks & Review Tracker");          // <-- Adjust tab name here
  const lastRow = sheet.getLastRow();
  const pdfUrl = pdfFile.getUrl();
  let targetRow = -1;

  for (let i = 2; i <= lastRow; i++) { // Assuming row 1 is header
    const richText = sheet.getRange(i, caseIdColumnIndex + 1).getRichTextValue();
    if (!richText) continue;

    const linkUrl = richText.getLinkUrl();

    if (linkUrl === caseId) {
      targetRow = i;
      break;
    }
  }

  // Change this column number if your form link is in another column (e.g. A = 1, B = 2)
  const formLinkColumn = 22;

  // Update the cell that previously contained the form link
  if (targetRow !== -1) {
    const displayText = `📄 PDF Generated Response`;
    const hyperlinkFormula = `=HYPERLINK("${pdfUrl}", "${displayText}")`;
    sheet.getRange(targetRow, formLinkColumn).setFormula(hyperlinkFormula);
  }

  // // 4. Log PDF link to sheet
  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // const lastRow = sheet.getLastRow();
  // sheet.getRange(lastRow, sheet.getLastColumn() + 1).setValue(pdfFile.getUrl());

  // // 5. Optional: Email the respondent a copy
  // MailApp.sendEmail({
  //   to: email,
  //   subject: `📝 Your Case Submission: ${caseId}`,
  //   body: `Thank you for your response. You can view your case summary here:\n\n${pdfFile.getUrl()}`,
  //   attachments: [pdfBlob],
  // });
}

/**
 * Serves the HTML file for the web app.
 * This is the main entry point when a user accesses the deployed web app URL.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('Post-Mortem Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Fetches and processes all data from the spreadsheet for the dashboard.
 * This single function prepares data for all cards and charts to optimize performance.
 * @returns {Object} A data object containing formatted information for cards and charts.
 */
/**
 * Fetches and processes all data from the spreadsheet for the dashboard.
 * Updated to split "Sources" by comma so combined values (e.g. "Google, Trustpilot") 
 * count towards their individual categories.
 *//**
 * Fetches and processes all data from the spreadsheet for the dashboard.
 */
function getDashboardData() {
  const SPREADSHEET_NAME = "PostMortem_Feedback & Issues";
  const SHEET_NAME = "Feedbacks & Review Tracker";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { error: "No data found." };
    
    // Get Values (Text/Numbers)
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const data = range.getValues();

    // Get Links specifically from Column A (Case ID)
    // We need RichTextValues to extract the URL from the HYPERLINK formula
    const linkData = sheet.getRange(2, 1, lastRow - 1, 1).getRichTextValues();

    // --- Data Processing ---
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    let totalCases = 0;
    let ongoingCases = 0;
    let casesThisMonth = 0;
    let reviewRemoved = 0;

    const categoryCounts = {};
    const marketCounts = {};
    const sourceCounts = {}; 
    const statusCounts = {};
    const marketCasesByMonth = {};
    const propertyCounts = {};
    
    const rawTableData = [];

    data.forEach((row, i) => {
      totalCases++;

      // Extract URL from the parallel linkData array
      const richText = linkData[i][0];
      const caseUrl = richText ? richText.getLinkUrl() : null;

      // A=0, B=1, C=2, F=5, G=6, H=7, I=8, L=11, P=15
      const caseIdText = row[0];
      const market = row[1] || 'Unknown';
      const sourceRaw = row[2] ? row[2].toString() : 'Unknown';
      const propertyId = row[5] || "Unknown";
      const reviewDateStr = row[6];
      const summary = row[7] || "";
      const category = row[8] || 'Uncategorized';
      const isRemoved = (row[11] && row[11].toString().toLowerCase() === "yes");
      const statusRaw = row[15] ? row[15].toString() : 'Unknown';
      
      // 1. Populate Table Data (Only what we need)
      rawTableData.push({
        id: caseIdText,
        url: caseUrl, // The extracted link
        summary: summary,
        // We still need these for the highlighting logic, even if not shown in table
        market: market,
        source: sourceRaw,
        category: category,
        status: statusRaw
      });

      // 2. Process Counts (Same as before)
      propertyCounts[propertyId] = (propertyCounts[propertyId] || 0) + 1;

      const statusLower = statusRaw.toLowerCase();
      if (statusLower === 'ongoing' || statusLower === 'open') ongoingCases++;
      statusCounts[statusRaw] = (statusCounts[statusRaw] || 0) + 1;

      if (isRemoved) reviewRemoved++;

      if (reviewDateStr && reviewDateStr instanceof Date) {
        if (reviewDateStr.getMonth() === currentMonth && reviewDateStr.getFullYear() === currentYear) {
          casesThisMonth++;
        }
        const monthYear = `${reviewDateStr.toLocaleString('default', { month: 'short' })} ${reviewDateStr.getFullYear()}`;
        if (!marketCasesByMonth[market]) marketCasesByMonth[market] = {};
        marketCasesByMonth[market][monthYear] = (marketCasesByMonth[market][monthYear] || 0) + 1;
      }
      
      categoryCounts[category] = (categoryCounts[category] || 0) + 1;
      marketCounts[market] = (marketCounts[market] || 0) + 1;

      const splitSources = sourceRaw.split(',').map(s => s.trim());
      splitSources.forEach(src => {
        if (src) sourceCounts[src] = (sourceCounts[src] || 0) + 1;
      });
    });

    // --- Prepare Data for Charts ---
    const repeatedProperties = Object.entries(propertyCounts)
      .filter(([pid, count]) => count >= 2)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);

    const repeatedPropertyData = {
      categories: repeatedProperties.map(item => item[0]),
      series: [{ name: "Repeats", data: repeatedProperties.map(item => item[1]) }]
    };

    const findTopItem = (counts) => Object.keys(counts).length ? Object.entries(counts).reduce((a, b) => a[1] > b[1] ? a : b)[0] : 'N/A';

    const allMonths = [...new Set(Object.values(marketCasesByMonth).flatMap(Object.keys))].sort((a,b) => new Date(a) - new Date(b));
    const marketSeries = Object.keys(marketCasesByMonth).map(mkt => {
      return {
        name: mkt,
        data: allMonths.map(month => marketCasesByMonth[mkt][month] || 0)
      };
    });

    return {
      cards: {
        totalCases: totalCases,
        ongoingCases: ongoingCases,
        casesThisMonth: casesThisMonth,
        topCategory: findTopItem(categoryCounts),
        topMarket: findTopItem(marketCounts),
        reviewRemoved: reviewRemoved
      },
      charts: {
        marketCasesByMonth: { series: marketSeries, categories: allMonths },
        casesBySource: { labels: Object.keys(sourceCounts), series: Object.values(sourceCounts) },
        casesByCategory: { categories: Object.keys(categoryCounts), series: [{ name: 'Cases', data: Object.values(categoryCounts) }] },
        caseStatus: { labels: Object.keys(statusCounts), series: Object.values(statusCounts) },
        repeatedProperties: repeatedPropertyData
      },
      rawTableData: rawTableData 
    };
  } catch (e) {
    return { error: e.message };
  }
}