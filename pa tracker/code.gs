function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('PA Non-Compliance Tracker')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fetches rows marked "not comply" for a specific country
function getDataForCountry(countryName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PA Charge report"); // Adjust to your exact tab name
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find column indexes (0-indexed)
  var resIdIdx = headers.indexOf("reservation_id");
  var countryIdx = headers.indexOf("country");
  var complianceIdx = headers.indexOf("PA Compliance");
  var reasonIdx = headers.indexOf("Manager Justification"); // Create this column header first!
  
  var filteredData = [];
  
  // Loop through rows (skip header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[complianceIdx] === "not comply" && row[countryIdx].toString().toLowerCase() === countryName.toLowerCase()) {
      filteredData.push({
        rowNum: i + 1, // Store exact spreadsheet row number for instant saving
        resId: row[resIdIdx],
        details: "Property: " + row[0] + " | Reason: " + row[3], // Customize what info they see
        currentReason: row[reasonIdx] || ""
      });
    }
  }
  return filteredData;
}

// Saves the reason back to the exact row in the heavy sheet
function saveReason(rowNum, reasonText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PA Charge report");
  var headers = sheet.getRowValues(1);
  var reasonColIdx = headers.indexOf("Manager Justification") + 1; // 1-indexed for sheets
  
  sheet.getRange(rowNum, reasonColIdx).setValue(reasonText);
  return "Saved successfully!";
}
