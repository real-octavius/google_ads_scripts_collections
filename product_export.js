function main() {
  var spreadsheetUrl = 'https://docs.google.com/spreadsheets/d/1zAuk6Zc4mI8USiQcQlA4U-dIUim81HMe18Tc8qvYhrs/edit?gid=639688247#gid=639688247';
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheet = spreadsheet.getSheetByName('Product Performance') || spreadsheet.insertSheet('Product Performance');
  
  // Get historical data before clearing current sheet
  var historicalData = getHistoricalData(spreadsheet);
  
  // Clear existing data
  sheet.clear();
  
  // Set headers (back to original structure)
  sheet.getRange(1, 1, 1, 11).setValues([
    ['Campaign Name', 'Product ID', 'Product Title', 'Conversions', 'Revenue', 'Clicks', 'Cost', 'Conv Rate', 'ROAS', 'Score', 'Rank']
  ]);
  
  try {
    // Build the query string
    var reportQuery = "SELECT " +
      "campaign.name, " +
      "segments.product_item_id, " +
      "segments.product_title, " +
      "metrics.conversions, " +
      "metrics.conversions_value, " +
      "metrics.clicks, " +
      "metrics.cost_micros " +
      "FROM shopping_performance_view " +
      "WHERE campaign.name = 'VL | Shopping | NB' " + // Replace with your campaign name
      "AND segments.date DURING LAST_30_DAYS " +
      "AND metrics.impressions > 0 " +
      "AND metrics.conversions > 0";
    
    Logger.log('Running query: ' + reportQuery);
    
    // Execute the report
    var report = AdsApp.report(reportQuery);
    var rows = report.rows();
    var data = [];
    var currentPerformanceData = {};
    
    while (rows.hasNext()) {
      var row = rows.next();
      
      var conversions = parseFloat(row['metrics.conversions']) || 0;
      var revenue = parseFloat(row['metrics.conversions_value']) / 1000000 || 0;
      var clicks = parseInt(row['metrics.clicks']) || 0;
      var cost = parseFloat(row['metrics.cost_micros']) / 1000000 || 0;
      var productId = row['segments.product_item_id'];
      var productTitle = row['segments.product_title'];
      
      var convRate = clicks > 0 ? conversions / clicks : 0;
      var roas = cost > 0 ? revenue / cost : 0;
      var score = conversions * 10 + convRate * 100 + roas * 5;
      
      // Store current performance data for history comparison
      currentPerformanceData[productId] = {
        title: productTitle,
        conversions: conversions,
        roas: roas
      };
      
      data.push([
        row['campaign.name'],
        productId,
        productTitle,
        conversions,
        revenue,
        clicks,
        cost,
        convRate,
        roas,
        score,
        0 // Rank will be added after sorting
      ]);
    }
    
    Logger.log('Found ' + data.length + ' products with conversions');
    
    if (data.length > 0) {
      // Sort by performance score (highest first)
      data.sort(function(a, b) { 
        return b[9] - a[9]; // Sort by score column
      });
      
      // Add rank numbers
      for (var i = 0; i < data.length; i++) {
        data[i][10] = i + 1;
      }
      
      // Write to main sheet
      sheet.getRange(2, 1, data.length, 11).setValues(data);
      
      // Format the main sheet
      formatProductPerformanceSheet(sheet, data.length);
      
      // Create/update performance history
      updatePerformanceHistory(spreadsheet, currentPerformanceData, historicalData);
      
      // Create top performers sheet (back to original structure)
      var topPerformersSheet = spreadsheet.getSheetByName('Top Performers') || spreadsheet.insertSheet('Top Performers');
      topPerformersSheet.clear();
      topPerformersSheet.getRange(1, 1, 1, 3).setValues([['Rank', 'Product ID', 'Product Title']]);
      
      var topCount = Math.min(50, data.length);
      var topData = [];
      
      for (var j = 0; j < topCount; j++) {
        topData.push([
          data[j][10], // Rank
          data[j][1],  // Product ID
          data[j][2]   // Product Title
        ]);
      }
      
      if (topData.length > 0) {
        topPerformersSheet.getRange(2, 1, topData.length, 3).setValues(topData);
        formatTopPerformersSheet(topPerformersSheet, topData.length);
      }
      
      Logger.log('Successfully exported ' + data.length + ' products');
      Logger.log('Top ' + topCount + ' performers saved to separate sheet');
      Logger.log('Performance history updated');
      
      // Send email notification
      sendEmailNotification(data.length, topCount);
    } else {
      Logger.log('No products found with conversions in the last 30 days');
      sendEmailNotification(0, 0);
    }
    
  } catch (error) {
    Logger.log('Error occurred: ' + error.toString());
  }
}

/**
* Retrieves historical performance data from the _Historical_Storage sheet
*/
function getHistoricalData(spreadsheet) {
try {
  var storageSheet = spreadsheet.getSheetByName('_Historical_Storage');
  if (!storageSheet) {
    Logger.log('No historical storage sheet found - this must be the first run');
    return {};
  }
  
  var historicalData = {};
  var dataRange = storageSheet.getDataRange();
  
  if (dataRange.getNumRows() <= 1) {
    Logger.log('No historical data found - this must be the first run');
    return historicalData;
  }
  
  var values = dataRange.getValues();
  
  // Skip header row and get the stored data from previous run
  for (var i = 1; i < values.length; i++) {
    var productId = values[i][0];
    var conversions = parseFloat(values[i][2]) || 0;
    var roas = parseFloat(values[i][3]) || 0;
    
    if (productId) {
      historicalData[productId] = {
        conversions: conversions,
        roas: roas
      };
    }
  }
  
  Logger.log('Retrieved historical data for ' + Object.keys(historicalData).length + ' products');
  return historicalData;
  
} catch (error) {
  Logger.log('Error retrieving historical data: ' + error.toString());
  return {};
}
}

/**
* Updates the Performance History sheet with current data and percentage changes
*/
function updatePerformanceHistory(spreadsheet, currentData, historicalData) {
try {
  var historySheet = spreadsheet.getSheetByName('Performance History') || spreadsheet.insertSheet('Performance History');
  
  // Clear existing data and set headers
  historySheet.clear();
  historySheet.getRange(1, 1, 1, 4).setValues([
    ['Product ID', 'Product Title', 'Conversions Change %', 'ROAS Change %']
  ]);
  
  var historyData = [];
  
  // Process each current product
  Object.keys(currentData).forEach(function(productId) {
    var current = currentData[productId];
    var historical = historicalData[productId];
    
    var conversionsChange = 'NEW';
    var roasChange = 'NEW';
    
    if (historical) {
      // Calculate percentage changes
      if (historical.conversions > 0) {
        var convChange = ((current.conversions - historical.conversions) / historical.conversions) * 100;
        conversionsChange = (convChange >= 0 ? '+' : '') + convChange.toFixed(1) + '%';
      } else if (current.conversions > 0) {
        conversionsChange = 'NEW CONVERSIONS';
      } else {
        conversionsChange = 'NO CHANGE';
      }
      
      if (historical.roas > 0) {
        var roasChangeVal = ((current.roas - historical.roas) / historical.roas) * 100;
        roasChange = (roasChangeVal >= 0 ? '+' : '') + roasChangeVal.toFixed(1) + '%';
      } else if (current.roas > 0) {
        roasChange = 'NEW ROAS';
      } else {
        roasChange = 'NO CHANGE';
      }
    }
    
    historyData.push([
      productId,
      current.title,
      conversionsChange,
      roasChange
    ]);
  });
  
  // Write the history data
  if (historyData.length > 0) {
    historySheet.getRange(2, 1, historyData.length, 4).setValues(historyData);
    formatPerformanceHistorySheet(historySheet, historyData);
    Logger.log('Updated Performance History with ' + historyData.length + ' products');
  }
  
  // Store current data for next run comparison in a separate sheet
  storeCurrentDataForNextRun(spreadsheet, currentData);
  
} catch (error) {
  Logger.log('Error updating performance history: ' + error.toString());
}
}

/**
* Stores current performance data for comparison in the next run
*/
function storeCurrentDataForNextRun(spreadsheet, currentData) {
try {
  var storageSheet = spreadsheet.getSheetByName('_Historical_Storage') || spreadsheet.insertSheet('_Historical_Storage');
  
  // Clear and set headers
  storageSheet.clear();
  storageSheet.getRange(1, 1, 1, 4).setValues([
    ['Product ID', 'Product Title', 'Conversions', 'ROAS']
  ]);
  
  var storageData = [];
  Object.keys(currentData).forEach(function(productId) {
    var data = currentData[productId];
    storageData.push([
      productId,
      data.title,
      data.conversions,
      data.roas
    ]);
  });
  
  if (storageData.length > 0) {
    storageSheet.getRange(2, 1, storageData.length, 4).setValues(storageData);
    Logger.log('Stored current data for next run comparison');
  }
  
  // Hide the storage sheet since it's just for internal use
  storageSheet.hideSheet();
  
} catch (error) {
  Logger.log('Error storing current data: ' + error.toString());
}
}

/**
* Formats the Product Performance sheet with colors and proper formatting
*/
function formatProductPerformanceSheet(sheet, dataRows) {
try {
  // Format headers
  var headerRange = sheet.getRange(1, 1, 1, 11);
  headerRange.setBackground('#4285F4')
            .setFontColor('#FFFFFF')
            .setFontWeight('bold')
            .setFontSize(11)
            .setHorizontalAlignment('center');
  
  if (dataRows > 0) {
    // Format data rows
    var dataRange = sheet.getRange(2, 1, dataRows, 11);
    dataRange.setFontSize(10)
             .setVerticalAlignment('middle');
    
    // Format specific columns
    // Conversions (column 4)
    sheet.getRange(2, 4, dataRows, 1).setNumberFormat('#,##0.00');
    
    // Revenue (column 5)
    sheet.getRange(2, 5, dataRows, 1).setNumberFormat('$#,##0.00');
    
    // Cost (column 7)
    sheet.getRange(2, 7, dataRows, 1).setNumberFormat('$#,##0.00');
    
    // Conv Rate (column 8)
    sheet.getRange(2, 8, dataRows, 1).setNumberFormat('0.00%');
    
    // ROAS (column 9)
    sheet.getRange(2, 9, dataRows, 1).setNumberFormat('#,##0.00');
    
    // Score (column 10)
    sheet.getRange(2, 10, dataRows, 1).setNumberFormat('#,##0.00');
    
    // Alternating row colors
    for (var i = 2; i <= dataRows + 1; i++) {
      var rowRange = sheet.getRange(i, 1, 1, 11);
      if (i % 2 === 0) {
        rowRange.setBackground('#F8F9FA');
      } else {
        rowRange.setBackground('#FFFFFF');
      }
    }
    
    // Add borders
    dataRange.setBorder(true, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 11);
  
  Logger.log('Product Performance sheet formatted successfully');
  
} catch (error) {
  Logger.log('Error formatting Product Performance sheet: ' + error.toString());
}
}

/**
* Formats the Performance History sheet with colors and conditional formatting
*/
function formatPerformanceHistorySheet(historySheet, historyData) {
try {
  // Format headers
  var headerRange = historySheet.getRange(1, 1, 1, 4);
  headerRange.setBackground('#34A853')
            .setFontColor('#FFFFFF')
            .setFontWeight('bold')
            .setFontSize(11)
            .setHorizontalAlignment('center');
  
  if (historyData.length > 0) {
    // Format data rows
    var dataRange = historySheet.getRange(2, 1, historyData.length, 4);
    dataRange.setFontSize(10)
             .setVerticalAlignment('middle');
    
    // Apply conditional formatting to percentage columns
    for (var i = 0; i < historyData.length; i++) {
      var rowNum = i + 2;
      
      // Format Conversions Change % (column 3)
      var convChangeCell = historySheet.getRange(rowNum, 3);
      var convChangeValue = historyData[i][2];
      formatPercentageCell(convChangeCell, convChangeValue);
      
      // Format ROAS Change % (column 4)
      var roasChangeCell = historySheet.getRange(rowNum, 4);
      var roasChangeValue = historyData[i][3];
      formatPercentageCell(roasChangeCell, roasChangeValue);
      
      // Alternating row background
      var rowRange = historySheet.getRange(rowNum, 1, 1, 4);
      if (i % 2 === 0) {
        if (rowRange.getBackground() === '#FFFFFF') {
          rowRange.setBackground('#F8F9FA');
        }
      }
    }
    
    // Add borders
    dataRange.setBorder(true, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // Auto-resize columns
  historySheet.autoResizeColumns(1, 4);
  
  Logger.log('Performance History sheet formatted successfully');
  
} catch (error) {
  Logger.log('Error formatting Performance History sheet: ' + error.toString());
}
}

/**
* Formats individual percentage cells with appropriate colors
*/
function formatPercentageCell(cell, value) {
try {
  if (typeof value === 'string') {
    if (value === 'NEW' || value === 'NEW CONVERSIONS' || value === 'NEW ROAS') {
      cell.setBackground('#FFF3CD')
          .setFontColor('#856404')
          .setFontWeight('bold');
    } else if (value.startsWith('+')) {
      cell.setBackground('#D4EDDA')
          .setFontColor('#155724')
          .setFontWeight('bold');
    } else if (value.startsWith('-')) {
      cell.setBackground('#F8D7DA')
          .setFontColor('#721C24')
          .setFontWeight('bold');
    } else if (value === '0%') {
      cell.setBackground('#E2E3E5')
          .setFontColor('#383D41');
    }
  }
} catch (error) {
  Logger.log('Error formatting percentage cell: ' + error.toString());
}
}

/**
* Formats the Top Performers sheet
*/
function formatTopPerformersSheet(sheet, dataRows) {
try {
  // Format headers
  var headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground('#FF6D00')
            .setFontColor('#FFFFFF')
            .setFontWeight('bold')
            .setFontSize(11)
            .setHorizontalAlignment('center');
  
  if (dataRows > 0) {
    // Format data rows
    var dataRange = sheet.getRange(2, 1, dataRows, 3);
    dataRange.setFontSize(10)
             .setVerticalAlignment('middle');
    
    // Highlight top 3 performers
    for (var i = 1; i <= Math.min(3, dataRows); i++) {
      var rowRange = sheet.getRange(i + 1, 1, 1, 3);
      switch (i) {
        case 1:
          rowRange.setBackground('#FFD700').setFontWeight('bold'); // Gold
          break;
        case 2:
          rowRange.setBackground('#C0C0C0').setFontWeight('bold'); // Silver
          break;
        case 3:
          rowRange.setBackground('#CD7F32').setFontWeight('bold'); // Bronze
          break;
      }
    }
    
    // Alternating colors for remaining rows
    for (var j = 4; j <= dataRows; j++) {
      var rowRange = sheet.getRange(j + 1, 1, 1, 3);
      if (j % 2 === 0) {
        rowRange.setBackground('#F8F9FA');
      }
    }
    
    // Add borders
    dataRange.setBorder(true, true, true, true, true, true, '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 3);
  
  Logger.log('Top Performers sheet formatted successfully');
  
} catch (error) {
  Logger.log('Error formatting Top Performers sheet: ' + error.toString());
}
}

/**
* Sends email notification when script runs
*/
function sendEmailNotification(productsCount, topPerformersCount) {
try {
  var email = 'lorenzo.filippini@queryclick.com';
  var subject = 'Google Ads Product Performance Script - Execution Complete';
  var timestamp = new Date().toLocaleString('en-GB', {timeZone: 'Europe/London'});
  
  var body = 'Hello Lorenzo,\n\n' +
             'The Google Ads Product Performance script has completed successfully.\n\n' +
             'Execution Details:\n' +
             '• Execution Time: ' + timestamp + '\n' +
             '• Products Processed: ' + productsCount + '\n' +
             '• Top Performers Identified: ' + topPerformersCount + '\n' +
             '• Performance History Updated: Yes\n\n';
  
  if (productsCount > 0) {
    body += 'The following sheets have been updated:\n' +
            '• Product Performance (main data with formatting)\n' +
            '• Performance History (with % changes and color coding)\n' +
            '• Top Performers (with rankings and highlighting)\n\n';
  } else {
    body += 'No products with conversions were found in the last 30 days.\n\n';
  }
  
  body += 'You can view the updated data in your Google Sheets.\n\n' +
          'Best regards,\n' +
          'Google Ads Automation Script';
  
  // Send the email using MailApp (available in Google Ads Scripts)
  MailApp.sendEmail(email, subject, body);
  
  Logger.log('Email notification sent to ' + email);
  
} catch (error) {
  Logger.log('Error sending email notification: ' + error.toString());
  // Don't throw error - email failure shouldn't stop the script
}
}