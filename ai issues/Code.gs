/**
 * Serves the HTML file for the Web App.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AI Issue Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches data from the 'November' tab.
 * Maps specific columns based on the prompt's requirements.
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('November');
  
  if (!sheet) {
    throw new Error('Sheet named "November" not found.');
  }

  const data = sheet.getDataRange().getDisplayValues();
  const headers = data.shift(); // Remove headers

  // Helper to convert Column Letter to Index (0-based)
  // A=0, O=14, P=15, AB=27, AD=29, AF=31, AG=32, AH=33, AK=36
  const col = (char) => {
    let result = 0;
    for (let i = 0; i < char.length; i++) {
      result *= 26;
      result += char.charCodeAt(i) - 64;
    }
    return result - 1;
  };

  const map = {
    id: col('A'),
    dueDate: col('C'),
    status: col('E'),
    propertyId: col('M'),
    country: col('O'),
    office: col('P'),
    platform: col('AB'),
    requestedBy: col('AD'), // Using AD for "Global GXO - Reviews Post-mortem" check
    category: col('AF'),
    subcategory: col('AG'),
    createdDate: col('AH'),
    rating: col('AK')
  };

  // Structure data for frontend
  const processedData = data.map(row => {
    return {
      id: row[map.id],
      dueDate: row[map.dueDate],
      status: row[map.status],
      propertyId: row[map.propertyId],
      country: row[map.country] ? row[map.country].trim() : "Unknown",
      office: row[map.office],
      platform: row[map.platform] || "Null",
      requestedBy: row[map.requestedBy],
      category: row[map.category],
      subcategory: row[map.subcategory],
      createdDate: row[map.createdDate], // String dd/mm/yyyy
      rating: parseFloat(row[map.rating]) || 0
    };
  });

  return JSON.stringify(processedData);
}