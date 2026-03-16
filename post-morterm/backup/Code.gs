// @ts-nocheck
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const column = range.getColumn();
  const row = range.getRow();

  if (column === 1) {
    const value = range.getValue();
    if (typeof value === "string" && value.startsWith("http")) {
      const displayText = "Case " + (row - 1);
      const formula = `=HYPERLINK("${value}", "${displayText}")`;
      sheet.getRange(row, column).setFormula(formula);
    }
  }
}

function onFormSubmit(e) {
  const responses = e.namedValues;
  const timestamp = responses["Timestamp"][0];
  const caseId = responses["Case ID"] ? responses["Case ID"][0] : "No ID Case"; 
  const caseIdColumnIndex = 0;
  const folderId = "1_gQPZoMSh7t9JHEQwFbURA48GxYL86ew"; 

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

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF);
  const pdfFile = DriveApp.getFolderById(folderId).createFile(pdfBlob);
  pdfFile.setName(`Response - ${caseId}.pdf`);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  const ss = SpreadsheetApp.openById("175WBoQpRxSaOBOQUSugM9OGA0JJfKYRTIkhUI6RiErw");
  const sheet = ss.getSheetByName("Feedbacks & Review Tracker");
  const lastRow = sheet.getLastRow();
  const pdfUrl = pdfFile.getUrl();
  let targetRow = -1;
  for (let i = 2; i <= lastRow; i++) { 
    const richText = sheet.getRange(i, caseIdColumnIndex + 1).getRichTextValue();
    if (!richText) continue;
    const linkUrl = richText.getLinkUrl();
    if (linkUrl === caseId) {
      targetRow = i;
      break;
    }
  }

  const formLinkColumn = 22;
  if (targetRow !== -1) {
    const displayText = `📄 PDF Generated Response`;
    const hyperlinkFormula = `=HYPERLINK("${pdfUrl}", "${displayText}")`;
    sheet.getRange(targetRow, formLinkColumn).setFormula(hyperlinkFormula);
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
    .setTitle('Post-Mortem Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// Added optional filter parameters for Month and Year
function getDashboardData(filterMonth, filterYear) {
  filterMonth = filterMonth || "All";
  filterYear = filterYear || "All";

  const SHEET_NAME = "Feedbacks & Review Tracker";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { error: "No data found." };
    
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const data = range.getValues();
    const linkData = sheet.getRange(2, 1, lastRow - 1, 1).getRichTextValues();
    
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    let totalCases = 0;
    let ongoingCases = 0;
    let reviewRemoved = 0;
    const categoryCounts = {};
    const marketCounts = {};
    const sourceCounts = {};
    const statusCounts = {};
    const marketCasesByMonth = {};
    const categoryCasesByMonth = {}; // <--- NEW: Object to track category frequency per month
    
    const propertyCounts = {};
    const propertyRawCounts = {};
    const propertyDedupeMap = {}; 

    const rawTableData = [];
    const availableYears = new Set();
    data.forEach(row => {
      const reviewDateStr = row[6];
      if (reviewDateStr && reviewDateStr instanceof Date) {
        availableYears.add(reviewDateStr.getFullYear());
      }
    });
    data.forEach((row, i) => {
      const reviewDateStr = row[6]; // Column G (Review Date)
      let rowMonth = null;
      let rowYear = null;
      const category = row[8] || 'Uncategorized'; // Moved category check up

      if (reviewDateStr && reviewDateStr instanceof Date) {
        rowMonth = reviewDateStr.getMonth() + 1; // 1 to 12
        rowYear = reviewDateStr.getFullYear();

        // <--- NEW: Category Heatmap Logic moved ABOVE the filters --->
        // This ensures the heatmap always gets all data regardless of the dashboard's current filter
        const monthYearHeatmap = `${reviewDateStr.toLocaleString('default', { month: 'short' })} ${reviewDateStr.getFullYear()}`;
        if (!categoryCasesByMonth[category]) categoryCasesByMonth[category] = {};
        categoryCasesByMonth[category][monthYearHeatmap] = (categoryCasesByMonth[category][monthYearHeatmap] || 0) + 1;
      }

      // --- FILTERS APPLY HERE ---
      if (filterYear !== "All" && String(rowYear) !== String(filterYear)) return;
      if (filterMonth !== "All" && String(rowMonth) !== String(filterMonth)) return;

      totalCases++;

      const richText = linkData[i][0];
      const caseUrl = richText ? richText.getLinkUrl() : null;

      const caseIdText = row[0];
      const market = row[1] || 'Unknown';
      const sourceRaw = row[2] ? row[2].toString() : 'Unknown';
      const affectedParty = row[3] ? row[3].toString() : "";
      const reservationId = row[4] ? row[4].toString() : "";
      
      const rawProp = row[5];
      const propertyId = (rawProp !== null && rawProp !== undefined && rawProp !== "") ? rawProp.toString().trim() : "Unknown";
      const summary = row[7] || "";
      // Category is already declared above
      const isRemoved = (row[11] && row[11].toString().toLowerCase() === "yes");
      const duplicateReviews = row[12] ? row[12].toString().toLowerCase().trim() : "";
      const statusRaw = row[15] ? row[15].toString() : 'Unknown';

      let displayReservationId = "";
      if (affectedParty.toLowerCase() === "guest") {
        displayReservationId = reservationId;
      } else if (affectedParty.toLowerCase() === "host") {
        displayReservationId = "Host";
      }

      let monthYearLabel = "Unknown Date";
      if (reviewDateStr && reviewDateStr instanceof Date) {
        monthYearLabel = `${reviewDateStr.toLocaleString('default', { month: 'long' })} ${reviewDateStr.getFullYear()}`;
      }

      rawTableData.push({
        id: caseIdText,
        url: caseUrl,
        summary: summary,
        market: market,
        source: sourceRaw,
        category: category,
        status: statusRaw,
        propertyId: propertyId,
        isRemoved: isRemoved,
        displayReservationId: displayReservationId,
        monthYear: monthYearLabel,
        reservationId: reservationId, 
        duplicateReviews: duplicateReviews 
      });
      if (propertyId !== "Unknown") {
        propertyRawCounts[propertyId] = (propertyRawCounts[propertyId] || 0) + 1;
        if (duplicateReviews !== "yes") {
          if (reservationId) {
             const uniqueKey = propertyId + "_" + reservationId;
             if (!propertyDedupeMap[uniqueKey]) {
                 propertyCounts[propertyId] = (propertyCounts[propertyId] || 0) + 1;
                 propertyDedupeMap[uniqueKey] = true;
             }
          } else {
             propertyCounts[propertyId] = (propertyCounts[propertyId] || 0) + 1;
          }
        }
      }

      const statusLower = statusRaw.toLowerCase();
      if (statusLower === 'ongoing' || statusLower === 'open') ongoingCases++;
      statusCounts[statusRaw] = (statusCounts[statusRaw] || 0) + 1;

      if (isRemoved) reviewRemoved++;
      if (reviewDateStr && reviewDateStr instanceof Date) {
        const monthYear = `${reviewDateStr.toLocaleString('default', { month: 'short' })} ${reviewDateStr.getFullYear()}`;
        // Existing Market Logic
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

    const repeatedProperties = Object.entries(propertyCounts)
      .filter(([pid, cleanCount]) => pid !== 'Unknown' && (propertyRawCounts[pid] || 0) >= 2 && cleanCount > 0)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
    const repeatedPropertyData = {
      categories: repeatedProperties.map(item => item[0]),
      series: [{
        name: "Repeats",
        data: repeatedProperties.map(item => item[1])
      }]
    };
    const findTopItem = (counts) => Object.keys(counts).length ? Object.entries(counts).reduce((a, b) => a[1] > b[1] ? a : b)[0] : 'N/A';
    // Sort all months chronologically (For filtered charts like Market Trend)
    const allMonths = [...new Set(Object.values(marketCasesByMonth).flatMap(Object.keys))].sort((a, b) => new Date(a) - new Date(b));
    // <--- NEW: Dedicated month array for Heatmap to show unfiltered range --->
    const allHeatmapMonths = [...new Set(Object.values(categoryCasesByMonth).flatMap(Object.keys))].sort((a, b) => new Date(a) - new Date(b));
    const marketSeries = Object.keys(marketCasesByMonth).map(mkt => {
      return {
        name: mkt,
        data: allMonths.map(month => marketCasesByMonth[mkt][month] || 0)
      };
    });
    // <--- NEW: Prepare Heatmap Series using unfiltered months array --->
    const heatmapSeries = Object.keys(categoryCasesByMonth).map(cat => {
      return {
        name: cat,
        data: allHeatmapMonths.map(month => categoryCasesByMonth[cat][month] || 0)
      };
    });
    const sortedCategories = Object.entries(categoryCounts).sort((a, b) => b[1] - a[1]);

    return {
      cards: {
        totalCases: totalCases,
        ongoingCases: ongoingCases,
        topCategory: findTopItem(categoryCounts),
        topMarket: findTopItem(marketCounts),
        reviewRemoved: reviewRemoved
      },
      charts: {
        marketCasesByMonth: { series: marketSeries, categories: allMonths },
        casesBySource: { labels: Object.keys(sourceCounts), series: Object.values(sourceCounts) },
        casesByCategory: { categories: sortedCategories.map(i => i[0]), series: [{ name: 'Cases', data: sortedCategories.map(i => i[1]) }] },
        caseStatus: { labels: Object.keys(statusCounts), series: Object.values(statusCounts) },
        repeatedProperties: repeatedPropertyData,
        // <--- NEW: Add Heatmap Data to return object mapped to unfiltered timeline --->
        categoryHeatmap: { series: heatmapSeries, categories: allHeatmapMonths }
      },
      rawTableData: rawTableData,
      availableYears: Array.from(availableYears).sort((a, b) => b - a) 
    };
  } catch (e) {
    return { error: e.message };
  }
}

// Function to fetch data for the new "Resolution Insights" page
function getResolutionInsightsData() {
  const SHEET_NAME = "Resolution Insights";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found. Please ensure it exists.`);
    // 1. Fetch Summary Data (B3:D4) - To capture the Grand Total as well
    const summaryRange = sheet.getRange("B3:D4");
    const summaryData = summaryRange.getDisplayValues();

    // 2. Fetch Definition Data (Row 6 onwards, Columns A to D)
    const lastRow = sheet.getLastRow();
    let defData = [];
    if (lastRow >= 6) {
      const defRange = sheet.getRange(`A6:D${lastRow}`);
      defData = defRange.getDisplayValues();
    }

    const processedDefs = [];
    let currentStrategy = "Uncategorized";
    for (let i = 0; i < defData.length; i++) {
      const row = defData[i];
      // Skip completely empty rows
      if (!row.join('').trim()) continue;
      // Skip the table header if it's there
      if (row[0].trim() === "Strategic Resolution" && row[1].trim() === "Issue") continue;
      // If Column A has text, update the current strategy (handles vertically merged cells)
      if (row[0].trim() !== "") {
        currentStrategy = row[0].trim();
      }
      
      processedDefs.push({
        strategy: currentStrategy,
        issue: row[1] ? row[1].trim() : "",
        solution: row[2] ? row[2].trim() : "",
        owner: row[3] ? row[3].trim() : ""
      });
    }

    return { summary: summaryData, definitions: processedDefs };
  } catch (e) {
    return { error: e.message };
  }
}