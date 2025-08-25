// Google Ads Script to retrieve keyword data for a specific campaign and write it to a Google Sheet.
// Replace the placeholders with your own values before running the script.
// To run the script, go to Google Ads > Tools and Settings > Bulk Actions > Scripts > + and paste the script.
// ** NOTE: Make sure to change your Google Sheet permission to edit before running the script, otherwise it will throw an error.


function main() {
    // Replace with your values here:
    let campaignName = "Campaign Name Here";
    let startDate = "20240510"; // Format: YYYYMMDD
    let endDate = "20240516"; // Format: YYYYMMDD
    let spreadsheetId = "Google Sheet ID Here";
    let sheetName = "Title of your sheet";
    
    // SQL query to retrieve demensions and metrics from the specified campaign.
    // If you want to add more dimensions or metrics, you can do so by adding them to the query.
    let report = AdsApp.report(
      "SELECT CampaignName, AdGroupName, Criteria, KeywordMatchType, Impressions, Clicks, AverageCpc, Cost, Conversions, ConversionValue  " +
      "FROM KEYWORDS_PERFORMANCE_REPORT " +
      "WHERE CampaignName = '" + campaignName + "' " +
      "DURING " + startDate + "," + endDate
    );
    
    let rows = report.rows();
    let data = [];

    while (rows.hasNext()) {
      let row = rows.next();
      data.push([
        row["CampaignName"],
        row["AdGroupName"],
        row["Criteria"],
        row["KeywordMatchType"],
        row["Impressions"],
        row["Clicks"],
        row["AverageCpc"],
        row["Cost"],
        row["Conversions"],
        row["ConversionValue"]
      ]);
    }
    
    let spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log("Sheet with name '" + sheetName + "' not found.");
      return;
    }
    
    sheet.clear();
    sheet.appendRow(["Campaign Name", "Ad Group Name", "Keyword", "Match Type", "Impressions", "Clicks", "CPC", "Cost", "Conversions", "Conversion Value"]);
    
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  }
  
