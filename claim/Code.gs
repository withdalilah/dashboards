/**
 * Serves the HTML Dashboard
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Claims Dashboard 2025')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Main function called by frontend
 */
function getDashboardData(filterMonth, filterYear) {
  var sheetName = "Claims - 2024"; // Update this to "Claims - 2025" when ready
  
  var ss = SpreadsheetApp.openById("1kWUFkVTJKWGmZguxNbWOmxDMoS4QkiUz7CzHK2b1sx4");
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) throw new Error("Sheet '" + sheetName + "' not found.");

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
  var rows = data.slice(1);

  // --- COLUMN MAPPING ---
  var colMap = {
    date: headers.findIndex(function(h) { return h.includes('date'); }),
    market: headers.findIndex(function(h) { return h.includes('market'); }),
    status: headers.findIndex(function(h) { return h.includes('status'); }),
    amount: headers.findIndex(function(h) { return h.includes('initial amount'); }), 
    payout: headers.findIndex(function(h) { return h.includes('final payout'); }),
    subject: headers.findIndex(function(h) { return h.includes('subject'); }),
    propId: headers.findIndex(function(h) { return (h.includes('property') || h.includes('prop')) && h.includes('id'); }),
    platform: headers.findIndex(function(h) { return (h.includes('platform') || h.includes('channel')) && !h.includes('id'); })
  };

  // Fallbacks
  if (colMap.propId === -1) colMap.propId = headers.findIndex(function(h) { return h.includes('property'); });

  var result = {
    kpi: { 
      won: { count: 0, value: 0 },
      lost: { count: 0, value: 0 },
      dropped: { count: 0, value: 0 },
      rejected: { count: 0, value: 0 } 
    },
    timeline: {}, 
    marketStatus: {}, 
    platformWinLoss: {}, 
    subjectStats: {}, 
    issueCounts: {}, // For Bar Chart (Filtered)
    rawCategoryTrend: {}, // For Trend Chart (Global)
    propertyIds: {} 
  };

  var parseMoney = function(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let str = val.toString().trim().replace(/[^0-9.,-]/g, '');
    if (!str) return 0;
    if (str.indexOf(',') > str.indexOf('.')) {
       str = str.replace(/\./g, '').replace(',', '.');
    } else {
       str = str.replace(/,/g, '');
    }
    return parseFloat(str) || 0;
  };

  // --- MAIN LOOP ---
  rows.forEach(function(row) {
    var rawDate = row[colMap.date];
    
    // 1. Date Parse
    var rowDateObj = null;
    if (rawDate instanceof Date) {
      rowDateObj = rawDate;
    } else if (typeof rawDate === 'string' && rawDate.includes('/')) {
      var parts = rawDate.split('/');
      if (parts.length === 3) rowDateObj = new Date(parts[2], parts[1] - 1, parts[0]);
    }
    if (!rowDateObj) return;

    var rowMonth = rowDateObj.getMonth() + 1; 
    var rowYear = rowDateObj.getFullYear();

    // 2. YEAR Filter (Always applies)
    if (rowYear != filterYear) return;

    // 3. Prepare Data
    var rawMarket = (row[colMap.market] || "Unknown").toString();
    var market = rawMarket.replace(/\s*\(.*?\)\s*/g, "").trim(); 
    if (market === "") market = "Unknown";
    var statusLower = (row[colMap.status] || "").toString().trim().toLowerCase();
    var subject = (row[colMap.subject] || "Unknown").toString().trim();

    // --- 4. GLOBAL CALCULATIONS (Before Month Filter) ---
    
    // A. Timeline (Total Cases)
    var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var timeKey = monthNames[rowDateObj.getMonth()] + " " + rowYear.toString().substring(2);
    if (!result.timeline[market]) result.timeline[market] = {};
    result.timeline[market][timeKey] = (result.timeline[market][timeKey] || 0) + 1;

    // B. Category Trend (For the Line Chart)
    // We store monthly counts for EVERY subject to find top ones later
    if (subject) {
      if (!result.rawCategoryTrend[subject]) result.rawCategoryTrend[subject] = new Array(12).fill(0);
      result.rawCategoryTrend[subject][rowDateObj.getMonth()]++; // Increment correct month index (0-11)
    }

    // C. Property IDs
    var propId = row[colMap.propId];
    if (propId !== "" && propId !== null && propId !== undefined) {
      if (!result.propertyIds[market]) result.propertyIds[market] = {};
      var pString = propId.toString();
      result.propertyIds[market][pString] = (result.propertyIds[market][pString] || 0) + 1;
    }

    // --- 5. MONTH FILTER BARRIER ---
    if (filterMonth !== "All" && rowMonth != filterMonth) return;

    // --- 6. FILTERED CALCULATIONS (KPIs, Charts) ---
    var amount = 0;
    if (statusLower === 'won') {
      amount = parseMoney(row[colMap.payout]); 
      result.kpi.won.count++; result.kpi.won.value += amount;
    } else if (statusLower === 'lost') {
      amount = parseMoney(row[colMap.amount]); 
      result.kpi.lost.count++; result.kpi.lost.value += amount;
    } else if (statusLower === 'dropped') {
      amount = parseMoney(row[colMap.amount]);
      result.kpi.dropped.count++; result.kpi.dropped.value += amount;
    } else if (statusLower === 'rejected') {
      amount = parseMoney(row[colMap.amount]);
      result.kpi.rejected.count++; result.kpi.rejected.value += amount;
    }

    // Status
    if (!result.marketStatus[market]) result.marketStatus[market] = { Won: 0, Lost: 0, Dropped: 0 };
    if (statusLower === 'won') result.marketStatus[market].Won++;
    else if (statusLower === 'lost') result.marketStatus[market].Lost++;
    else if (statusLower === 'dropped') result.marketStatus[market].Dropped++;

    // Platform
    var platform = (row[colMap.platform] || "Other").toString().trim();
    if (!result.platformWinLoss[platform]) result.platformWinLoss[platform] = { Won: 0, Lost: 0 };
    if (statusLower === 'won') result.platformWinLoss[platform].Won++;
    else if (statusLower === 'lost') result.platformWinLoss[platform].Lost++;

    // Subject Stats (Won vs Lost)
    if (subject) {
      if (!result.subjectStats[market]) result.subjectStats[market] = { won: {}, lost: {} };
      if (statusLower === 'won') result.subjectStats[market].won[subject] = (result.subjectStats[market].won[subject] || 0) + 1;
      else if (statusLower === 'lost') result.subjectStats[market].lost[subject] = (result.subjectStats[market].lost[subject] || 0) + 1;
      
      // Issue Counts (Filtered)
      result.issueCounts[subject] = (result.issueCounts[subject] || 0) + 1;
    }
  });

  // --- POST PROCESSING ---

  // 1. Process Category Trend (Get Top 5 Global)
  var allTrends = [];
  for (var subj in result.rawCategoryTrend) {
    var total = result.rawCategoryTrend[subj].reduce((a, b) => a + b, 0);
    allTrends.push({ name: subj, data: result.rawCategoryTrend[subj], total: total });
  }
  // Sort by total annual count desc and take top 5
  allTrends.sort((a,b) => b.total - a.total);
  result.categoryTrend = allTrends.slice(0, 5);
  delete result.rawCategoryTrend;

  // 2. Process Issue Counts (Filtered Top 10)
  var sortedIssues = [];
  for (var key in result.issueCounts) {
    sortedIssues.push({ category: key, count: result.issueCounts[key] });
  }
  sortedIssues.sort((a,b) => b.count - a.count);
  result.issueCounts = sortedIssues.slice(0, 10);

  // 3. Process Subjects (Win/Loss)
  var finalSubjects = [];
  for (var mkt in result.subjectStats) {
    var getTop = function(obj) {
      var topKey = ""; var topVal = 0;
      for (var k in obj) { if (obj[k] > topVal) { topVal = obj[k]; topKey = k; } }
      return { subject: topKey, count: topVal };
    };
    var topWon = getTop(result.subjectStats[mkt].won);
    var topLost = getTop(result.subjectStats[mkt].lost);
    if (topWon.count > 0 || topLost.count > 0) {
      finalSubjects.push({
        market: mkt,
        wonSubject: topWon.subject, wonCount: topWon.count,
        lostSubject: topLost.subject, lostCount: topLost.count
      });
    }
  }
  result.topSubjects = finalSubjects;

  // 4. Process Property IDs
  var topProps = [];
  for (var mkt in result.propertyIds) {
    var propsArr = [];
    for (var id in result.propertyIds[mkt]) {
      propsArr.push({ id: id, count: result.propertyIds[mkt][id] });
    }
    propsArr.sort((a,b) => b.count - a.count);
    topProps.push({ market: mkt, items: propsArr.slice(0, 3) });
  }
  result.topProperties = topProps;

  delete result.subjectStats;
  delete result.propertyIds; 

  return JSON.stringify(result);
}