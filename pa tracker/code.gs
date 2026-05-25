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
  
  // Find all required columns dynamically
  var resIdIdx = headers.indexOf("reservation_id");
  var countryIdx = headers.indexOf("country");
  var complianceIdx = headers.indexOf("PA Compliance");
  var monthIdx = headers.indexOf("Month");
  var noteIdx = headers.indexOf("Note from OS");
  var resultIdx = headers.indexOf("result"); // <--- ADDED: Looks for Column O
  
  // Best practice: Find Property ID and Reason dynamically too, just in case columns shift
  var propIdIdx = headers.indexOf("property_id"); 
  var reasonIdx = headers.indexOf("reason"); 
  
  var filteredData = [];
  var totalReservationsInMonth = 0;
  var monthNum = parseInt(selectedMonth, 10);
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    if (row[countryIdx].toString().toLowerCase() === countryName.toLowerCase() && parseInt(row[monthIdx], 10) === monthNum) {
      
      totalReservationsInMonth++; 
      
      if (row[complianceIdx] === "not comply") {
        
        // Grab values securely
        var prop = propIdIdx > -1 ? row[propIdIdx] : row[0];
        var rsn = reasonIdx > -1 ? row[reasonIdx] : row[3];
        var resResult = resultIdx > -1 ? row[resultIdx] : "N/A"; // <--- Grab the result
        
        filteredData.push({
          rowNum: i + 1,
          resId: row[resIdIdx],
          // ADDED: The result text is now appended to the details display
          details: "Property: " + prop + " | Reason: " + rsn + " | Result: " + resResult, 
          currentReason: row[noteIdx] || "" 
        });
      }
    }
  }
  
  return {
    rows: filteredData,
    totalCount: totalReservationsInMonth
  };
}

function saveReason(rowNum, reasonText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PA Charge report");
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; 
  var noteColIdx = headers.indexOf("Note from OS") + 1; 
  
  if (noteColIdx === 0) { noteColIdx = 25; } 
  
  sheet.getRange(rowNum, noteColIdx).setValue(reasonText);
  
  SpreadsheetApp.flush(); 
  
  return "Saved successfully!";
}