function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('PA Non-Compliance Tracker')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAvailableMonths() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PA Charge report");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var monthIdx = headers.indexOf("Month"); 
  
  if (monthIdx === -1) return [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]; 
  
  var months = [];
  for (var i = 1; i < data.length; i++) {
    var mVal = data[i][monthIdx];
    if (mVal && months.indexOf(mVal) === -1 && !isNaN(mVal)) {
      months.push(mVal);
    }
  }
  return months.sort(function(a, b){return a - b});
}

function getDataForCountryAndMonth(countryName, selectedMonth) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PA Charge report");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var resIdIdx = headers.indexOf("reservation_id");
  var countryIdx = headers.indexOf("country");
  var complianceIdx = headers.indexOf("PA Compliance");
  var monthIdx = headers.indexOf("Month");
  var noteIdx = headers.indexOf("Note from OS"); 
  
  var filteredData = [];
  var totalReservationsInMonth = 0;
  var monthNum = parseInt(selectedMonth, 10);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Check if it belongs to the selected country and month
    if (row[countryIdx].toString().toLowerCase() === countryName.toLowerCase() && parseInt(row[monthIdx], 10) === monthNum) {
      
      totalReservationsInMonth++; // Count total reservations for this market/month
      
      // If it is ALSO non-compliant, add it to our list
      if (row[complianceIdx] === "not comply") {
        filteredData.push({
          rowNum: i + 1,
          resId: row[resIdIdx],
          details: "Property: " + row[0] + " | Reason: " + row[3], 
          currentReason: row[noteIdx] || "" 
        });
      }
    }
  }
  
  // Return BOTH the list of bad rows AND the total count
  return {
    rows: filteredData,
    totalCount: totalReservationsInMonth
  };
}

function saveReason(rowNum, reasonText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PA Charge report");
  
  // FIXED: Correct way to get headers in Apps Script
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; 
  var noteColIdx = headers.indexOf("Note from OS") + 1; 
  
  if (noteColIdx === 0) { noteColIdx = 25; } // 25 = Column Y
  
  sheet.getRange(rowNum, noteColIdx).setValue(reasonText);
  
  // Forces Google to write to the sheet immediately before returning success
  SpreadsheetApp.flush(); 
  
  return "Saved successfully!";
}