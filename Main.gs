// Initialization
function main() {
  initializeSpreadsheet(); 
  importAccPerformance();  
  importActivityFeed();    
  importTokenStats();      
  createDashboard();       
}

function initializeSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfNotExists(ss, 'Dashboard');
  deleteSheet1();
  createSheetIfNotExists(ss, 'Acc Performance');
  createSheetIfNotExists(ss, 'Activity Feed');
  createSheetIfNotExists(ss, 'Historical Token Stats');
  
}

function createSheetIfNotExists(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ss.insertSheet(sheetName);
  }
}

function deleteSheet1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");
  
  // Check if the sheet exists before attempting to delete
  if (sheet) {
    ss.deleteSheet(sheet);
  }
}

// Imports Mango account performance related data (balances, PnL etc.)
function importAccPerformance() {
  var url = "https://api.mngo.cloud/data/v4/stats/performance_account?mango-account=#INSERT_YOUR_MANGO_ACOUNT_ID_HERE#";
  var response = UrlFetchApp.fetch(url);
  var json = JSON.parse(response.getContentText());

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Acc Performance";
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  var headers = ["Date & Time (UTC)"];
  var data = [];

  for (var key in json) {
    if (json.hasOwnProperty(key)) {
      var date = new Date(key);  // Updated to directly use the timestamp from the JSON
      var formattedDate = Utilities.formatDate(date, "GMT", "yyyy-MM-dd HH:mm");
      var row = [formattedDate];
      var obj = json[key];

      if (headers.length === 1) {
        for (var prop in obj) {
          if (obj.hasOwnProperty(prop) && prop !== 'timestamp') {  // Exclude the original timestamp from the data
            headers.push(prop);
          }
        }
        data.push(headers);
      }

      for (var i = 1; i < headers.length; i++) {
        row.push(parseFloat(obj[headers[i]]).toFixed(2));
      }
      data.push(row);
    }
  }

  var renamedHeaders = ["Date & Time (UTC)", "Acc Balance", "PnL", "Spot Value", "Perp Value", "Transfer Balance", "Deposit Interest", "Borrow Interest", "Spot Volume"];
  data[0] = renamedHeaders;

  sheet.clear();
  var range = sheet.getRange(1, 1, data.length, headers.length);
  range.setValues(data);

  // Apply Solarized dark mode color scheme to entire sheet
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground("#073642").setFontColor("#839496").setFontFamily("VT323");
  
  // Updated the header text color to yellow
  sheet.getRange(1, 1, 1, headers.length).setFontColor("#b58900").setFontWeight("bold");

  // Additional formatting
  sheet.setFrozenRows(1);
  
  // Autoresize Columns
  sheet.autoResizeColumns(1, 9);
}

// Imports Mango account activity related data (trades, deposits/withdrawals, etc)
function importActivityFeed() {
  var url = "https://api.mngo.cloud/data/v4/stats/activity-feed?mango-account=#INSERT_YOUR_MANGO_ACOUNT_ID_HERE#";
  
  // Fetch the API data
  var response = UrlFetchApp.fetch(url);
  var data = JSON.parse(response.getContentText());
  
  // Open or create the "Activity Feed" sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Activity Feed");
  if(!sheet) {
    sheet = ss.insertSheet("Activity Feed");
  } else {
    sheet.clear();
  }
  
  // Function to transform date-time format
  function transformDateTime(dateTime) {
    var parts = dateTime.split("T");
    var date = parts[0];
    var time = parts[1].split(":").slice(0, 2).join(":");
    return date + " " + time;
  }
  
  // If data is present, populate the sheet
  if(data && data.length > 0) {
    // Predefined columns
    var predefinedHeaders = ["Date & Time (UTC)", "activity_type"];
    
    // Collecting all unique keys from "activity_details" for headers
    var detailKeys = new Set();
    data.forEach(function(item) {
      Object.keys(item.activity_details).forEach(function(key) {
        if(key !== "block_datetime") {  // Avoiding the second block_datetime
          detailKeys.add(key);
        }
      });
    });
    
    function toTitleCase(str) {
      // Placeholder for bracket content
      var placeholders = [];

      // Replace bracket content with placeholders
      var newStr = str.replace(/\((.*?)\)/g, function(match) {
        placeholders.push(match);
        return `(${placeholders.length - 1})`;
      });

      // Convert the string outside of brackets to title case
      newStr = newStr.replace(/_/g, ' ').replace(/\w\S*/g, function(word) {
        return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
      });

      // Restore original bracket content from placeholders
      newStr = newStr.replace(/\((\d+)\)/g, function(match, p1) {
        return placeholders[parseInt(p1, 10)];
      });

      return newStr;
    }
  

    var headers = predefinedHeaders.concat(Array.from(detailKeys));
    // Cleaning up the header titles
    var cleanedHeaders = headers.map(function(header) {
      if(header === "wallet_pk") return "Wallet";
      if(header === "usd_equivalent") return "USD Value";
      if(header === "swap_in_price_usd") return "Swap In Price USD";
      if(header === "swap_out_price_usd") return "Swap Out Price USD";
      if(header === "maker_order_id") return "Maker Order ID";
      if(header === "taker_order_id") return "Taker Order ID";
      if(header === "taker_client_order_id") return "Taker Client Order ID";
      if(header === "order_id") return "Order ID";
      return toTitleCase(header);
    });
    sheet.appendRow(cleanedHeaders);
    
  // Populating rows
  for(var i = 0; i < data.length; i++) {
    var row = headers.map(function(header) {
      var value;
      if(header === "Date & Time (UTC)") {
        value = transformDateTime(data[i]["block_datetime"]);
      } else if(predefinedHeaders.includes(header)) {
        value = data[i][header];
      } else {
        value = data[i].activity_details[header];
      }

      if (value === null) return "null";
      if (value === false) return "false";
      if (value === 0) return "0";
      return value || "";
    });
    sheet.appendRow(row);
  }
  } else {
    Logger.log("No data found or error fetching data.");
  }
  // Apply Solarized dark mode color scheme to entire sheet
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground("#073642").setFontColor("#839496").setFontFamily("VT323").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  
  // Update the header text color to yellow
  sheet.getRange(1, 1, 1, headers.length).setFontColor("#b58900").setFontWeight("bold");

  // Freeze first row
  sheet.setFrozenRows(1);
  
  // Resize Columns
  sheet.setColumnWidths(1, 12, 104);
  sheet.setColumnWidth(1, 98);  // Date & Time (UTC)
  sheet.setColumnWidth(2, 82);  // Activity type
  sheet.setColumnWidth(6, 44);  // Symbol
  sheet.setColumnWidth(7, 66);  // Quantity
  sheet.setColumnWidth(8, 66);  // USD Value
  sheet.setColumnWidth(11, 34); // Bid
  sheet.setColumnWidth(12, 55); // Maker
  sheet.setColumnWidth(16, 50); // Fee Tier
  sheet.setColumnWidth(17, 87); // Instruction Num
  sheet.setColumnWidth(18, 34); // Size
  sheet.setColumnWidth(19, 34); // Price
  sheet.setColumnWidth(20, 28); // Side
  sheet.setColumnWidth(21, 55); // Fee Cost
  sheet.setColumnWidth(23, 66); // Base Symbol
  sheet.setColumnWidth(24, 71); // Quote Symbol
  sheet.setColumnWidth(25, 55); // Slot
  sheet.setColumnWidth(26, 82); // Maker Order ID
  sheet.setColumnWidth(29, 82); // Taker Order ID
  sheet.setColumnWidth(31, 87); // Taker Fee
  sheet.setColumnWidth(32, 60); // Taker Side
  sheet.setColumnWidth(34, 71); // Market Index
  sheet.setColumnWidth(35, 44); // Sequence Number
  sheet.setColumnWidth(38, 82); // Swap in Symbol
  sheet.setColumnWidth(39, 82); // Swap In Amount
  sheet.setColumnWidth(41, 87); // Swap Out Symbol
  sheet.setColumnWidth(42, 87); // Swap Out Amount
  sheet.setColumnWidth(44, 66); // Loan
  sheet.setColumnWidth(45, 114);// Loan Origination Fee
}

// Imports general token related market data
function importTokenStats() {
  var url = "https://api.mngo.cloud/data/v4/token-historical-stats?mango-group=78b8f4cGCwmZ9ysPFMWLaLTkkaYnUjwMJYStWe5RTSSX";
  var response = UrlFetchApp.fetch(url);
  var jsonData = JSON.parse(response.getContentText());

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Historical Token Stats";
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  var headers = [];
  var data = [];

  // Map for renaming headers
  var headerRenameMap = {
    'token_index': 'Token Index',
    'symbol': 'Symbol',
    'price': 'Price',
    'stable_price': 'Stable Price',
    'total_deposits': 'Total Deposits',
    'total_borrows': 'Total Borrows',
    'collected_fees': 'Collected Fees',
    'deposit_apr': 'Deposit APR',
    'borrow_apr': 'Borrow APR',
    'deposit_rate': 'Deposit Rate',
    'borrow_rate': 'Borrow Rate'
  };

  jsonData.forEach(function(item, index) {
    var date = new Date(item.date_hour);
    var formattedDate = Utilities.formatDate(date, "GMT", "yyyy-MM-dd HH:mm");
    var row = [formattedDate];

    // If it's the first record, populate headers
    if (index === 0) {
      headers.push("Date & Time (UTC+2)");
      for (var prop in item) {
        if (prop !== 'date_hour' && prop !== 'mango_group') {
          headers.push(headerRenameMap[prop] || prop);
        }
      }
      data.push(headers);
    }

    // Populate the data
    for (var i = 1; i < headers.length; i++) {
      var originalHeader = Object.keys(headerRenameMap).find(key => headerRenameMap[key] === headers[i]) || headers[i];
      row.push(parseFloat(item[originalHeader]) || item[originalHeader]);
    }
    data.push(row);
  });

  sheet.clear();
  var range = sheet.getRange(1, 1, data.length, headers.length);
  range.setValues(data);

  // Apply Solarized dark mode color scheme to entire sheet
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setBackground("#073642").setFontColor("#839496").setFontFamily("VT323");
  
  // Update the header text color to yellow
  sheet.getRange(1, 1, 1, headers.length).setFontColor("#b58900").setFontWeight("bold");

  // Freeze first row
  sheet.setFrozenRows(1);
  
  // Resize Columns
  sheet.setColumnWidths(1, 12, 104);
}

// Dashboard Section

var ss = SpreadsheetApp.getActiveSpreadsheet();
var dashboard = ss.getSheetByName('Dashboard');
var dataSheetAP = ss.getSheetByName('Acc Performance');
var dataSheetHTS = ss.getSheetByName('Historical Token Stats');

// Creates the entire Sheet
function createDashboard() {
  if (dashboard) {
    ss.deleteSheet(dashboard);
  }
  dashboard = ss.insertSheet('Dashboard', 0);
  
  formatDashboard();
  createAccPerfChart();
  createTokenStatsChart();
  createAccPerfChart2();
  createTrigger();
}

// Dashboard Formatting
function formatDashboard() {
  // Apply Solarized dark mode color scheme to the dashboard
  dashboard.getRange(1, 1, dashboard.getMaxRows(), dashboard.getMaxColumns()).setBackground('#073642').setFontColor('#839496');

  // Adds formatting to cellrange B2 to G2
  dashboard.getRange('B2:G2').setBackground('#fdf6e3').setFontColor('#cb4b16').setFontSize(12).setFontFamily('VT323').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dashboard.getRange('F2:G2').merge();
  dashboard.getRange('C2').setHorizontalAlignment('left');
  dashboard.getRange('E2').setHorizontalAlignment('left');

  // Writes Start and End Date into the cell range and protects it
  dashboard.getRange('B2').setValue('Start Date');
  dashboard.getRange('D2').setValue('End Date');
  var cell = dashboard.getRange('B2');
  var dataValidationRule = SpreadsheetApp.newDataValidation()
  .requireTextEqualTo('Start Date')  // The cell value must be equal to "Start Date"
  .setAllowInvalid(false)  // This will show a warning when a user tries to change the cell content
  .setHelpText('You cannot change this cell!')  // Custom warning message
  .build();
  cell.setDataValidation(dataValidationRule);
  var cell = dashboard.getRange('D2');
  var dataValidationRule = SpreadsheetApp.newDataValidation()
  .requireTextEqualTo('End Date')  // The cell value must be equal to "End Date"
  .setAllowInvalid(false)  // This will show a warning when a user tries to change the cell content
  .setHelpText('You cannot change this cell!')  // Custom warning message
  .build();
  cell.setDataValidation(dataValidationRule);

  // Creates date dropdowns for Chart 1
  var lastRow = dataSheetAP.getLastRow();
  var dateRange = dataSheetAP.getRange(2, 1, lastRow-1);  // Adjusted to exclude the header row
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(dateRange).build();
  dashboard.getRange('C2').setDataValidation(rule);
  dashboard.getRange('E2').setDataValidation(rule);

  // Adds formatting to cell range B23 to G23
  dashboard.getRange('B23:G23').setBackground('#fdf6e3').setFontColor('#cb4b16').setFontFamily('VT323').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(12);

  // Writes Token A and Token B into the cell range and protects it
  dashboard.getRange('B23').setValue('Token A');
  dashboard.getRange('D23').setValue('Token B');
  var cell = dashboard.getRange('B23');
  var dataValidationRule = SpreadsheetApp.newDataValidation()
  .requireTextEqualTo('Token A')  // The cell value must be equal to "Token A"
  .setAllowInvalid(false)  // This will show a warning when a user tries to change the cell content
  .setHelpText('You cannot change this cell!')  // Custom warning message
  .build();
  cell.setDataValidation(dataValidationRule);
  var cell = dashboard.getRange('D23');
  var dataValidationRule = SpreadsheetApp.newDataValidation()
  .requireTextEqualTo('Token B')  // The cell value must be equal to "Token B"
  .setAllowInvalid(false)  // This will show a warning when a user tries to change the cell content
  .setHelpText('You cannot change this cell!')  // Custom warning message
  .build();
  cell.setDataValidation(dataValidationRule);

  // Adds formatting to cellrange I2 to N2
  dashboard.getRange('I2:N2').setBackground('#fdf6e3').setFontColor('#cb4b16').setFontSize(12).setFontFamily('VT323').setHorizontalAlignment('center').setVerticalAlignment('middle');
  dashboard.getRange('M2:N2').merge();
  dashboard.getRange('J2').setHorizontalAlignment('left');
  dashboard.getRange('L2').setHorizontalAlignment('left');

  // Writes Start and End Date into the cell range and protects it
  dashboard.getRange('I2').setValue('Start Date');
  dashboard.getRange('K2').setValue('End Date');
  var cell = dashboard.getRange('I2');
  var dataValidationRule = SpreadsheetApp.newDataValidation()
  .requireTextEqualTo('Start Date')  // The cell value must be equal to "Start Date"
  .setAllowInvalid(false)  // This will show a warning when a user tries to change the cell content
  .setHelpText('You cannot change this cell!')  // Custom warning message
  .build();
  cell.setDataValidation(dataValidationRule);
  var cell = dashboard.getRange('K2');
  var dataValidationRule = SpreadsheetApp.newDataValidation()
  .requireTextEqualTo('End Date')  // The cell value must be equal to "End Date"
  .setAllowInvalid(false)  // This will show a warning when a user tries to change the cell content
  .setHelpText('You cannot change this cell!')  // Custom warning message
  .build();
  cell.setDataValidation(dataValidationRule);

  // Creates date dropdowns for chart 3
  var lastRow = dataSheetAP.getLastRow();
  var dateRange2 = dataSheetAP.getRange(2, 1, lastRow-1);  // Adjusted to exclude the header row
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(dateRange2).build();
  dashboard.getRange('J2').setDataValidation(rule);
  dashboard.getRange('L2').setDataValidation(rule);

  // Resize column width of column A & H 
  dashboard.setColumnWidth(1, 21);
  dashboard.setColumnWidth(8, 21);

}

// Creates 1st Chart
function createAccPerfChart() {
  // Fetch Spot Value, Perp Value, Dates and put it into FetchedData AP sheet (hidden)
  var lastRow = dataSheetAP.getLastRow();
  var fetchedDataAPSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FetchedData AP");
  if (!fetchedDataAPSheet) {
      fetchedDataAPSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("FetchedData AP");
  } else {
      fetchedDataAPSheet.clear();
  }
  fetchedDataAPSheet.getRange(1, 1).setValue('Date');
  fetchedDataAPSheet.getRange(1, 2).setValue('Spot Value');
  fetchedDataAPSheet.getRange(1, 3).setValue('Perp Value');
  var dateValues = dataSheetAP.getRange(2, 1, lastRow - 1).getValues();
  var spotValues = dataSheetAP.getRange(2, 4, lastRow - 1).getValues();
  var perpValues = dataSheetAP.getRange(2, 5, lastRow - 1).getValues();
  fetchedDataAPSheet.getRange(2, 1, lastRow - 1, 1).setValues(dateValues).setNumberFormat("MM/dd");
  fetchedDataAPSheet.getRange(2, 2, lastRow - 1, 1).setValues(spotValues);
  fetchedDataAPSheet.getRange(2, 3, lastRow - 1, 1).setValues(perpValues);

  // Calculate min and max of Spot Value & Perp Value
  var rawMinSpotValue = Math.min.apply(null, spotValues.map(function(row) { return row[0]; }));
  var rawMaxSpotValue = Math.max.apply(null, spotValues.map(function(row) { return row[0]; }));
  var rawMinPerpValue = Math.min.apply(null, perpValues.map(function(row) { return row[0]; }));
  var rawMaxPerpValue = Math.max.apply(null, perpValues.map(function(row) { return row[0]; }));

  // Calculate the range between max and min for Spot Value
  var spotRange = rawMaxSpotValue - rawMinSpotValue;
  var spotPadding = 0.10 * spotRange;

  // Adjusted min and max for Spot Value
  var minSpotValue = rawMinSpotValue - spotPadding;
  var maxSpotValue = rawMaxSpotValue + spotPadding;

  // Calculate the range between max and min for Perp Value
  var perpRange = rawMaxPerpValue - rawMinPerpValue;
  var perpPadding = 0.10 * perpRange;

  // Adjusted min and max for Perp Value
  var minPerpValue = rawMinPerpValue - perpPadding;
  var maxPerpValue = rawMaxPerpValue + perpPadding;

  // Defining the range for Date, Spot Value & Perp Value using the fetched data
  var spotValueRange = fetchedDataAPSheet.getRange(1, 2, lastRow);
  var perpValueRange = fetchedDataAPSheet.getRange(1, 3, lastRow);
  var dateRange = fetchedDataAPSheet.getRange(1, 1, lastRow);
  
  // Chart Creation
  // Define common styles and settings
  var commonTextStyle = {color: '#839496', fontname: 'VT323'};
  var commonGridlines = {color: 'none'};
  var backgroundColor = '#fdf6e3';
  var chartColors = ['#859900', '#DC322F'];

  // Builds the chart
  var chartBuilder = dashboard.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dateRange)
    .addRange(spotValueRange)
    .addRange(perpValueRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('useFirstColumnAsDomain', true)
    .setOption('backgroundColor', backgroundColor)
    .setOption('colors', chartColors)
    .setOption('series', {
        0: {targetAxisIndex: 0, labelInLegend: 'Spot Value'},
        1: {targetAxisIndex: 1, labelInLegend: 'Perp Value'}
    })
    .setOption('vAxes', {
        0: {
            title: 'Spot Value',
            titleTextStyle: commonTextStyle,
            textStyle: commonTextStyle,
            gridlines: commonGridlines,
            viewWindow: {min: minSpotValue, max: maxSpotValue}
        },        
        1: {
            title: 'Perp Value',
            titleTextStyle: commonTextStyle,
            textStyle: commonTextStyle,
            gridlines: commonGridlines,
            viewWindow: {min: minPerpValue, max: maxPerpValue}
        }        
    })
    .setOption('hAxis.textStyle', commonTextStyle)
    .setOption('legend', {textStyle: commonTextStyle})
    .setPosition(4, 2, 0, 0)
    .setOption('hAxis', {
        gridlines: commonGridlines,
        minorGridlines: {count: 0},
        textStyle: commonTextStyle
    });
  // Build and insert the chart
  dashboard.insertChart(chartBuilder.build());
  fetchedDataAPSheet.hideSheet();
}

// Creates 2nd Chart
function createTokenStatsChart() {
  var lastRow = dataSheetHTS.getLastRow();

  // Filter data to get rows for SOL and MNGO
  var allData = dataSheetHTS.getRange(1, 1, lastRow, 4).getValues();
  var solData = allData.filter(row => row[2] === 'SOL');
  var mngoData = allData.filter(row => row[2] === 'MNGO');
  
  // Extract Dates and Prices
  var solDates = solData.map(row => [row[0]]);
  var solPrices = solData.map(row => [row[3]]);
  var mngoPrices = mngoData.map(row => [row[3]]);

  // Calculate min and max prices for SOL and MNGO
  var solMin = Math.min(...solPrices.map(row => row[0]));
  var solMax = Math.max(...solPrices.map(row => row[0]));
  var mngoMin = Math.min(...mngoPrices.map(row => row[0]));
  var mngoMax = Math.max(...mngoPrices.map(row => row[0]));

  // Adjust these values by 10%
  var solMinAdjusted = solMin * 0.9;
  var solMaxAdjusted = solMax * 1.1;
  var mngoMinAdjusted = mngoMin * 0.9;
  var mngoMaxAdjusted = mngoMax * 1.1;
  
  // Define common styles and settings
  var commonTextStyle = {color: '#839496', fontname: 'VT323'};
  var commonGridlines = {color: 'none'};
  var backgroundColor = '#fdf6e3';  // Updated color
  var chartColors = ['#859900', '#DC322F'];

  // Check if the "FetchedData" sheet exists
  var fetchedDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FetchedData HTS");

  // If it doesn't exist, create it. If it does, clear its contents
  if (!fetchedDataSheet) {
    fetchedDataSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("FetchedData HTS");
  } else {
    fetchedDataSheet.clear();
  }

  // sets headers
  fetchedDataSheet.getRange(1, 1).setValue('Date');
  fetchedDataSheet.getRange(1, 2).setValue('SOL Price');
  fetchedDataSheet.getRange(1, 3).setValue('MNGO Price');

  // Hide the "FetchedData HTS" sheet
  fetchedDataSheet.hideSheet();

  // Set the values of the new sheet to the fetched data
  fetchedDataSheet.getRange(2, 1, solDates.length, 1).setValues(solDates).setNumberFormat("MM/dd");
  fetchedDataSheet.getRange(2, 2, solPrices.length, 1).setValues(solPrices);
  fetchedDataSheet.getRange(2, 3, mngoPrices.length, 1).setValues(mngoPrices);

  // Builds the chart
  var chartBuilder = dashboard.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(fetchedDataSheet.getRange(1, 1, solDates.length, 1)) // Date range
    .addRange(fetchedDataSheet.getRange(1, 2, solPrices.length, 1)) // SOL Prices
    .addRange(fetchedDataSheet.getRange(1, 3, mngoPrices.length, 1)) // MNGO Prices
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('useFirstColumnAsDomain', true)
    .setOption('backgroundColor', backgroundColor)
    .setPosition(25, 2, 0, 0)
    .setOption('colors', chartColors)
    .setOption('series', {
        0: {targetAxisIndex: 0, labelInLegend: 'SOL'},
        1: {targetAxisIndex: 1, labelInLegend: 'MNGO'}
    })
    .setOption('vAxes', {
      0: {
          title: 'SOL Price',
          titleTextStyle: commonTextStyle,
          textStyle: commonTextStyle,
          gridlines: commonGridlines,
          viewWindow: {
            min: solMinAdjusted,
            max: solMaxAdjusted
          }
      },
      1: {
          title: 'MNGO Price',
          titleTextStyle: commonTextStyle,
          textStyle: commonTextStyle,
          gridlines: commonGridlines,
          viewWindow: {
            min: mngoMinAdjusted,
            max: mngoMaxAdjusted
          }
      }
    })
    .setOption('hAxis', {
      gridlines: commonGridlines,
      minorGridlines: {count: 0},
      textStyle: commonTextStyle,
      majorGridlines: {count: 6},
      ticksPosition: 'inside',
      ticksLength: 6
    });
    

  // Build and insert the chart
  dashboard.insertChart(chartBuilder.build());

}

// Creates 3rd Chart
function createAccPerfChart2() {
  // Fetch Date & Time (UTC), Acc Balance, and PnL, and put them into FetchedData AP2 sheet (hidden)
  var lastRow = dataSheetAP.getLastRow();
  var fetchedDataAP2Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FetchedData AP2");
  if (!fetchedDataAP2Sheet) {
      fetchedDataAP2Sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("FetchedData AP2");
  } else {
      fetchedDataAP2Sheet.clear();
  }
  fetchedDataAP2Sheet.getRange(1, 1).setValue('Date');
  fetchedDataAP2Sheet.getRange(1, 2).setValue('Acc Balance');
  fetchedDataAP2Sheet.getRange(1, 3).setValue('PnL');
  
  var dateValues = dataSheetAP.getRange(2, 1, lastRow - 1).getValues();
  var accBalanceValues = dataSheetAP.getRange(2, 2, lastRow - 1).getValues();
  var pnlValues = dataSheetAP.getRange(2, 3, lastRow - 1).getValues();
  
  fetchedDataAP2Sheet.getRange(2, 1, lastRow - 1, 1).setValues(dateValues).setNumberFormat("MM/dd");
  fetchedDataAP2Sheet.getRange(2, 2, lastRow - 1, 1).setValues(accBalanceValues);
  fetchedDataAP2Sheet.getRange(2, 3, lastRow - 1, 1).setValues(pnlValues);
  
  // Calculate min and max of Acc Balance & PnL for padding
  var rawMinAccBalance = Math.min.apply(null, accBalanceValues.map(function(row) { return row[0]; }));
  var rawMaxAccBalance = Math.max.apply(null, accBalanceValues.map(function(row) { return row[0]; }));
  var accBalanceValueRange = rawMaxAccBalance - rawMinAccBalance; // Renamed variable
  var accBalancePadding = 0.10 * accBalanceValueRange; // Updated reference
  var minAccBalance = rawMinAccBalance - accBalancePadding;
  var maxAccBalance = rawMaxAccBalance + accBalancePadding;

  var rawMinPnL = Math.min.apply(null, pnlValues.map(function(row) { return row[0]; }));
  var rawMaxPnL = Math.max.apply(null, pnlValues.map(function(row) { return row[0]; }));
  var pnlValueRange = rawMaxPnL - rawMinPnL; // Renamed variable
  var pnlPadding = 0.10 * pnlValueRange; // Updated reference
  var minPnL = rawMinPnL - pnlPadding;
  var maxPnL = rawMaxPnL + pnlPadding;

  // Define common styles and settings
  var commonTextStyle = {color: '#839496', fontname: 'VT323'};
  var commonGridlines = {color: 'none'};
  var backgroundColor = '#fdf6e3';
  var chartColors = ['#859900', '#DC322F'];

  // Builds the chart
  var chartBuilder = dashboard.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(fetchedDataAP2Sheet.getRange(1, 1, lastRow, 1)) // Date range
    .addRange(fetchedDataAP2Sheet.getRange(1, 2, lastRow, 1)) // Acc Balance
    .addRange(fetchedDataAP2Sheet.getRange(1, 3, lastRow, 1)) // PnL
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setTransposeRowsAndColumns(false)
    .setNumHeaders(1)
    .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
    .setOption('useFirstColumnAsDomain', true)
    .setOption('backgroundColor', backgroundColor)
    .setOption('colors', chartColors)
    .setOption('series', {
        0: {targetAxisIndex: 0, labelInLegend: 'Acc Balance'},
        1: {targetAxisIndex: 1, labelInLegend: 'PnL'}
    })
    .setOption('vAxes', {
        0: {
            title: 'Acc Balance',
            titleTextStyle: commonTextStyle,
            textStyle: commonTextStyle,
            gridlines: commonGridlines,
            viewWindow: {min: minAccBalance, max: maxAccBalance}
        },
        1: {
            title: 'PnL',
            titleTextStyle: commonTextStyle,
            textStyle: commonTextStyle,
            gridlines: commonGridlines,
            viewWindow: {min: minPnL, max: maxPnL}
        }
    })
    .setOption('hAxis.textStyle', commonTextStyle)
    .setOption('legend', {textStyle: commonTextStyle})
    .setPosition(4, 9, 0, 0)
    .setOption('hAxis', {
        gridlines: commonGridlines,
        minorGridlines: {count: 0},
        textStyle: commonTextStyle
    });

  // Build and insert the chart
  dashboard.insertChart(chartBuilder.build());
  fetchedDataAP2Sheet.hideSheet();
}

// Creates onEdit Trigger
function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  
  // Iterate through all triggers
  for (var i = 0; i < triggers.length; i++) {
    // If an onEdit trigger is found, delete it
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
      break;  // Exit the loop once the trigger is found and deleted
    }
  }
  
  // Add a new onEdit trigger
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
}

// Trigger Logic for Both Chart 1 and Chart 3 Dropdowns
function onEdit(e) {
  // Check if this is the third execution
  var scriptProperties = PropertiesService.getScriptProperties();
  var executionCount = scriptProperties.getProperty('executionCount') || 0;
  executionCount = parseInt(executionCount, 10) + 1;
  Logger.log('Execution count: ' + executionCount); // Log the current count
  
  if (executionCount < 3) {
    // Exit if this is not the third execution
    scriptProperties.setProperty('executionCount', executionCount);
    Logger.log('Exiting early, not the third execution.');
    return;
  } else {
    // Reset the counter if this is the third execution
    scriptProperties.setProperty('executionCount', 0);
    Logger.log('This is the third execution, proceeding with updates.');
  }

  var range = e.range;
  var col = range.getColumn();
  var row = range.getRow();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');

  // Logic for Chart 1 Dropdown (C2 or E2)
  if (row === 2 && (col === 3 || col === 5)) {
    var startDateCell = dashboard.getRange('C2');
    var endDateCell = dashboard.getRange('E2');
    var messageCell = dashboard.getRange('F2');
    Logger.log('Handling date selection for Chart 1.');
    handleDateSelection(startDateCell, endDateCell, messageCell, updateChartData);
  }

  // Logic for Chart 3 Date Inputs (J2 or L2)
  if (row === 2 && (col === 10 || col === 12)) {
    var startDateCell = dashboard.getRange('J2');
    var endDateCell = dashboard.getRange('L2');
    var messageCell = dashboard.getRange('M2');
    Logger.log('Handling date selection for Chart 3.');
    handleDateSelection(startDateCell, endDateCell, messageCell, updateChartData2);
  }
}

// Helper function to handle date selection logic common to both charts
function handleDateSelection(startDateCell, endDateCell, messageCell, updateChartFunction) {
  Logger.log("startDateCell A1 Notation: " + startDateCell.getA1Notation());
  Logger.log("endDateCell A1 Notation: " + endDateCell.getA1Notation());
  var startDate = startDateCell.getValue();
  var endDate = endDateCell.getValue();

  Logger.log("Start Date: " + startDate);
  Logger.log("End Date: " + endDate);

  if (startDate && !endDate) {
    messageCell.setValue('Select End Date');
  } else if (!startDate && endDate) {
    messageCell.setValue('Select Start Date');
  } else if (startDate && endDate) {
    if (startDate < endDate) {
      Logger.log("Calling update function: " + updateChartFunction.name);
      updateChartFunction(startDate, endDate);
      messageCell.setValue('');  // Clear the error message
    } else {
      messageCell.setValue('Error: End Date Before Start Date');
    }
  }
}

// Updates 1st Chart with Date Range Selection
function updateChartData(startDate, endDate) {
  Logger.log("Entered updateChartData with startDate: " + startDate + ", endDate: " + endDate);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');

  // fetch data & put it into "FetchedData AP"
  var lastRow = dataSheetAP.getLastRow();
  var fetchedDataAPSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FetchedData AP");
  if (!fetchedDataAPSheet) {
      fetchedDataAPSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("FetchedData AP");
  } else {
      fetchedDataAPSheet.clear();
  }
  fetchedDataAPSheet.getRange(1, 1).setValue('Date');
  fetchedDataAPSheet.getRange(1, 2).setValue('Spot Value');
  fetchedDataAPSheet.getRange(1, 3).setValue('Perp Value');
  var dateValues = dataSheetAP.getRange(2, 1, lastRow - 1).getValues();
  var spotValues = dataSheetAP.getRange(2, 4, lastRow - 1).getValues();
  var perpValues = dataSheetAP.getRange(2, 5, lastRow - 1).getValues();
  fetchedDataAPSheet.getRange(2, 1, lastRow - 1, 1).setValues(dateValues).setNumberFormat("MM/dd");
  fetchedDataAPSheet.getRange(2, 2, lastRow - 1, 1).setValues(spotValues);
  fetchedDataAPSheet.getRange(2, 3, lastRow - 1, 1).setValues(perpValues);

  // Error handling for date selection
  if (startDate && endDate && startDate < endDate) {
    Logger.log("Valid date range selected");
    
    // Fetch data from the main data source
    var dateValues = dataSheetAP.getRange(2, 1, lastRow - 1).getValues();

    // Assume the data is in descending order
    var startRow, endRow;
    var foundStartRow = false;

    // Loop to find the start and end rows based on the selected dates
    for (var i = 0; i < dateValues.length; i++) {
      var dateTime = new Date(dateValues[i][0]);

      if (!foundStartRow && dateTime <= endDate) {
        startRow = i + 2;  // Adjusted for the header and 0-based index
        foundStartRow = true;
      }

      if (foundStartRow && dateTime < startDate) {
        endRow = i + 1;  // Previous row, as the current one is past the startDate
        break;
      }
    }

    // If startRow wasn't set, then startDate was never less than any date
    if (!foundStartRow) {
      startRow = lastRow;  // Assuming lastRow is the default when no matching start date is found
    }

    // If endRow wasn't set, then endDate was never greater than any date
    if (!endRow) {
      endRow = 2;  // Assuming 2 is the default when no matching end date is found
    }

    // Swap startRow and endRow if startRow is greater than endRow
    if (startRow > endRow) {
      var temp = startRow;
      startRow = endRow;
      endRow = temp;
    }

    if (startRow && endRow) {
      Logger.log("Start row: " + startRow + ", End row: " + endRow);
      // Define dateRange, spotValueRange, and perpValueRange here
      var dateRange = fetchedDataAPSheet.getRange(startRow, 1, endRow - startRow + 1);
      var spotValueRange = fetchedDataAPSheet.getRange(startRow, 2, endRow - startRow + 1);
      var perpValueRange = fetchedDataAPSheet.getRange(startRow, 3, endRow - startRow + 1);


      // Remove the old chart
      var charts = dashboard.getCharts();
      charts.forEach(function(chart) {
          var position = chart.getContainerInfo();
          var row = position.getAnchorRow();
          var col = position.getAnchorColumn();
          
          // Check if the chart is anchored at the desired position
          if (row === 4 && col === 2) {
              dashboard.removeChart(chart);
              Logger.log("Deleted a chart at position: Row " + row + ", Column " + col);
          }
      });


      // Fetch the data for calculating min and max values
      var spotValues = dataSheetAP.getRange(startRow, 4, endRow - startRow + 1).getValues();
      var perpValues = dataSheetAP.getRange(startRow, 5, endRow - startRow + 1).getValues();

      // Calculate the range between max and min for Spot Value
      var rawMinSpotValue = Math.min.apply(null, spotValues.map(function(row) { return row[0]; }));
      var rawMaxSpotValue = Math.max.apply(null, spotValues.map(function(row) { return row[0]; }));
      var spotRange = rawMaxSpotValue - rawMinSpotValue;
      var spotPadding = 0.10 * spotRange;
      var minSpotValue = rawMinSpotValue - spotPadding;
      var maxSpotValue = rawMaxSpotValue + spotPadding;

      // Calculate the range between max and min for Perp Value
      var rawMinPerpValue = Math.min.apply(null, perpValues.map(function(row) { return row[0]; }));
      var rawMaxPerpValue = Math.max.apply(null, perpValues.map(function(row) { return row[0]; }));
      var perpRange = rawMaxPerpValue - rawMinPerpValue;
      var perpPadding = 0.10 * perpRange;
      var minPerpValue = rawMinPerpValue - perpPadding;
      var maxPerpValue = rawMaxPerpValue + perpPadding;
      Logger.log("Min Spot Value: " + minSpotValue + ", Max Spot Value: " + maxSpotValue);
      Logger.log("Min Perp Value: " + minPerpValue + ", Max Perp Value: " + maxPerpValue);

      
      // Chart Creation
      // Define common styles and settings
      var commonTextStyle = {color: '#839496'};
      var commonGridlines = {color: 'none'};
      var backgroundColor = '#fdf6e3';
      var chartColors = ['#859900', '#DC322F'];

      // Builds the chart
      var chartBuilder = dashboard.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(dateRange)
        .addRange(spotValueRange)
        .addRange(perpValueRange)
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
        .setTransposeRowsAndColumns(false)
        .setNumHeaders(1)
        .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
        .setOption('useFirstColumnAsDomain', true)
        .setOption('backgroundColor', backgroundColor)
        .setOption('colors', chartColors)
        .setOption('series', {
            0: {targetAxisIndex: 0, labelInLegend: 'Spot Value'},
            1: {targetAxisIndex: 1, labelInLegend: 'Perp Value'}
        })
        .setOption('vAxes', {
            0: {
                title: 'Spot Value',
                titleTextStyle: commonTextStyle,
                textStyle: commonTextStyle,
                gridlines: commonGridlines,
                viewWindow: {min: minSpotValue, max: maxSpotValue}
            },
            1: {
                title: 'Perp Value',
                titleTextStyle: commonTextStyle,
                textStyle: commonTextStyle,
                gridlines: commonGridlines,
                viewWindow: {min: minPerpValue, max: maxPerpValue}
            }
        })
        .setOption('hAxis.textStyle', commonTextStyle)
        .setOption('legend', {textStyle: commonTextStyle})
        .setPosition(4, 2, 0, 0)
        .setOption('hAxis', {
            gridlines: commonGridlines,
            minorGridlines: {count: 0},
            textStyle: commonTextStyle
        });
      // Build and insert the chart
      Logger.log("Inserting new chart");
      dashboard.insertChart(chartBuilder.build());

    } else {
      Logger.log("No data available for the selected date range");
      dashboard.getRange('F2').setValue("No data available for the selected date range");
    }
  } else {
    Logger.log("Error: Invalid date range");
    dashboard.getRange('F2').setValue("Error: Invalid date range");
  }
}

// Updates 3rd Chart with Date Range Selection
function updateChartData2(startDate, endDate) {
  Logger.log('Entered updateChartData2 with startDate: ' + startDate + ', endDate: ' + endDate);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');

  // Extract data from the main data source and populate the temporary hidden sheet for the third chart
  var lastRow = dataSheetAP.getLastRow();
  var fetchedDataAP2Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FetchedData AP2");
  if (!fetchedDataAP2Sheet) {
      fetchedDataAP2Sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("FetchedData AP2");
  } else {
      fetchedDataAP2Sheet.clear();
  }
  fetchedDataAP2Sheet.getRange(1, 1).setValue('Date');
  fetchedDataAP2Sheet.getRange(1, 2).setValue('Acc Balance');
  fetchedDataAP2Sheet.getRange(1, 3).setValue('PnL');

  // Fetch data from the main data source
  var dateValues = dataSheetAP.getRange(2, 1, lastRow - 1).getValues();
  var accBalanceValues = dataSheetAP.getRange(2, 2, lastRow - 1).getValues();
  var pnlValues = dataSheetAP.getRange(2, 3, lastRow - 1).getValues();

  // Populate the temporary hidden sheet with the fetched data
  fetchedDataAP2Sheet.getRange(2, 1, lastRow - 1, 1).setValues(dateValues).setNumberFormat("MM/dd");
  fetchedDataAP2Sheet.getRange(2, 2, lastRow - 1, 1).setValues(accBalanceValues);
  fetchedDataAP2Sheet.getRange(2, 3, lastRow - 1, 1).setValues(pnlValues);

  // Error handling for date selection
  if (startDate && endDate && startDate < endDate) {
    // Get all dates from the data sheet
    var allDates = dataSheetAP.getRange(2, 1, lastRow - 1).getValues();
    var startRow, endRow;

    // Assume the data is in descending order
    var foundStartRow = false;
    for (var i = 0; i < allDates.length; i++) {
      var dateTime = new Date(allDates[i][0]);
      Logger.log('Row ' + (i+2) + ' Date: ' + dateTime);

      // Set startRow on the first date that is less than or equal to endDate
      if (!foundStartRow && dateTime <= endDate) {
        startRow = i + 2;
        foundStartRow = true;
        Logger.log('Start row set to: ' + startRow);
      }

      // Assuming we are iterating from the most recent date backwards,
      // set endRow when we reach a date that is less than the startDate
      if (foundStartRow && dateTime < startDate) {
        endRow = i + 1; // We want the previous row, as the current one is already past the startDate
        Logger.log('End row set to: ' + endRow);
        break; // Break the loop after finding endRow
      }
    }


    // After the loop, check if startRow was ever set, if not, then it means startDate was never less than any date, hence set it to lastRow
    if (!startRow) {
      startRow = endRow;
    }

    Logger.log('Final Start row: ' + startRow + ', End row: ' + endRow);



    // Check if the startRow is greater than the endRow, which shouldn't happen in descending order
    if (startRow > endRow) {
      Logger.log('Error: startRow is greater than endRow');
      // You might want to handle this case, e.g., show an error message or reset the rows
    }

    if (startRow && endRow) {
      Logger.log('Final Start row: ' + startRow + ', End row: ' + endRow);
    } else {
      Logger.log('Start row or end row not found');
    }

    // Swap startRow and endRow if data is in descending order
    if (startRow > endRow) {
      var temp = startRow;
      startRow = endRow;
      endRow = temp;
    }

    if (startRow && endRow) {
      // Define dateRange, accBalanceRange, and pnlRange here
      var dateRange = fetchedDataAP2Sheet.getRange(startRow, 1, endRow - startRow + 1, 1);
      var accBalanceChartRange = fetchedDataAP2Sheet.getRange(startRow, 2, endRow - startRow + 1, 1);
      var pnlChartRange = fetchedDataAP2Sheet.getRange(startRow, 3, endRow - startRow + 1, 1);
      Logger.log("Date range for chart: " + dateRange.getA1Notation());
      Logger.log("Acc Balance range for chart: " + accBalanceChartRange.getA1Notation()); // Updated variable name
      Logger.log("PnL range for chart: " + pnlChartRange.getA1Notation()); // Updated variable name
      Logger.log("dateRange type: " + (typeof dateRange));
      Logger.log("accBalanceChartRange type: " + (typeof accBalanceChartRange));
      Logger.log("pnlChartRange type: " + (typeof pnlChartRange));


      // Remove the old chart
      var charts = dashboard.getCharts();
      charts.forEach(function(chart) {
          var position = chart.getContainerInfo();
          var row = position.getAnchorRow();
          var col = position.getAnchorColumn();
          
          // Check if the chart is anchored at the desired position
          if (row === 4 && col === 9) {
              dashboard.removeChart(chart);
              Logger.log("Deleted a chart at position: Row " + row + ", Column " + col);
          }
      });


      // Fetch the data for calculating min and max values
      var accBalanceValues = dataSheetAP.getRange(startRow, 2, endRow - startRow + 1).getValues();  // Adjusted column index for Acc Balance
      var pnlValues = dataSheetAP.getRange(startRow, 3, endRow - startRow + 1).getValues();  // Adjusted column index for PnL

      // Calculate the range between max and min for Acc Balance
      var rawMinAccBalance = Math.min.apply(null, accBalanceValues.map(function(row) { return row[0]; }));
      var rawMaxAccBalance = Math.max.apply(null, accBalanceValues.map(function(row) { return row[0]; }));
      var accBalanceRange = rawMaxAccBalance - rawMinAccBalance;
      var accBalancePadding = 0.10 * accBalanceRange;
      var minAccBalance = rawMinAccBalance - accBalancePadding;
      var maxAccBalance = rawMaxAccBalance + accBalancePadding;

      // Calculate the range between max and min for PnL
      var rawMinPnL = Math.min.apply(null, pnlValues.map(function(row) { return row[0]; }));
      var rawMaxPnL = Math.max.apply(null, pnlValues.map(function(row) { return row[0]; }));
      var pnlRange = rawMaxPnL - rawMinPnL;
      var pnlPadding = 0.10 * pnlRange;
      var minPnL = rawMinPnL - pnlPadding;
      var maxPnL = rawMaxPnL + pnlPadding;

      // Define common styles and settings
      var commonTextStyle = {color: '#839496', fontname: 'VT323'};
      var commonGridlines = {color: 'none'};
      var backgroundColor = '#fdf6e3';
      var chartColors = ['#859900', '#DC322F'];

      // Builds the chart
      var chartBuilder = dashboard.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(dateRange)
        .addRange(accBalanceChartRange)
        .addRange(pnlChartRange)
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
        .setTransposeRowsAndColumns(false)
        .setNumHeaders(1)
        .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
        .setOption('useFirstColumnAsDomain', true)
        .setOption('backgroundColor', backgroundColor)
        .setOption('colors', chartColors)
        .setOption('series', {
            0: {targetAxisIndex: 0, labelInLegend: 'Acc Balance'},
            1: {targetAxisIndex: 1, labelInLegend: 'PnL'}
        })
        .setOption('vAxes', {
            0: {
                title: 'Acc Balance',
                titleTextStyle: commonTextStyle,
                textStyle: commonTextStyle,
                gridlines: commonGridlines,
                viewWindow: {min: minAccBalance, max: maxAccBalance}
            },
            1: {
                title: 'PnL',
                titleTextStyle: commonTextStyle,
                textStyle: commonTextStyle,
                gridlines: commonGridlines,
                viewWindow: {min: minPnL, max: maxPnL}
            }
        })
        .setOption('hAxis.textStyle', commonTextStyle)
        .setOption('legend', {textStyle: commonTextStyle})
        .setPosition(4, 9, 0, 0)
        .setOption('hAxis', {
            gridlines: commonGridlines,
            minorGridlines: {count: 0},
            textStyle: commonTextStyle
        });
      
      Logger.log('Min Acc Balance: ' + minAccBalance + ', Max Acc Balance: ' + maxAccBalance);
      Logger.log('Min PnL: ' + minPnL + ', Max PnL: ' + maxPnL);
      Logger.log('Inserting new chart');
      // Build and insert the chart
      dashboard.insertChart(chartBuilder.build());
      fetchedDataAP2Sheet.hideSheet();

    } else {
      dashboard.getRange('M2').setValue("No data available for the selected date range");  // Adjusted the cell range for third chart's error message
    }
  } else {
    dashboard.getRange('M2').setValue("Error: Invalid date range");  // Adjusted the cell range for third chart's error message
  }
}

