// Automatically moves all of the information on the Complete sheet to an archive - runs on the 1st of each month at midnight
function moveAndClearCompleteSheet() {
  // Get the current date
  let dateOnly = today.toISOString().split('T')[0]; // Only want the date without the time

  // Open the source and destination spreadsheets
  let sourceSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19rV2sw-JVJaYHYmG0xC5FrpqE9uGjKYKOdBOiwvBxI4/edit#gid=240309260");
  let destinationSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1NtwzZ3Q9BqjNHmfY1E7cC5u3-PmZhSj-5nUtyTRbDC8/edit#gid=0");

  // Get the data from the source sheet
  let sourceSheet = sourceSpreadsheet.getSheetByName("Complete");
  let data = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();

  // Create the destination sheet
  destinationSheet = destinationSpreadsheet.insertSheet(dateOnly);
  
  // Insert a row for column headers before appending the data
  let columnHeaders = ["L", "Category", "Brand", "SubCat", "Item", "Stock", "Sold L.M.", "List Date", "Order/Transfer Date", "Arrival Date", "Quantity", "Employee (Receiving)", "Employee (Sending)", "Notes", "T/O"];
  destinationSheet.getRange(1, 1, 1, columnHeaders.length).setValues([columnHeaders]);

  // Apply formatting to column headers
  let headerRange = destinationSheet.getRange(1, 1, 1, columnHeaders.length);
  headerRange.setBackground('#0b5394'); // Background color
  headerRange.setFontColor('#FFFFFF'); // Font color
  headerRange.setFontWeight('bold');  // Font weight
  destinationSheet.setFrozenRows(1);
  destinationSheet.getRange(1, 1, 1, columnHeaders.length).createFilter();

  // Append the data to the destination sheet starting from row 2
  destinationSheet.getRange(destinationSheet.getLastRow() + 1, 1, data.length - 1, data[0].length).setValues(data.slice(1));

  // Apply conditional formatting rule
  let ordersHighlightRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('O')
    .setBackground('#b1e0d3')  // Set the background color to #b1e0d3
    .setRanges([destinationSheet.getRange(2, 1, destinationSheet.getLastRow() - 1, destinationSheet.getLastColumn())])  // Specify the range to apply the rule
    .build();

  let transfersHighlightRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('T')
    .setBackground('#c9b8e6')  // Set the background color to #c9b8e6
    .setRanges([destinationSheet.getRange(2, 1, destinationSheet.getLastRow() - 1, destinationSheet.getLastColumn())])  // Specify the range to apply the rule
    .build();

  // Remove existing rules (optional)
  destinationSheet.clearConditionalFormatRules();

  // Apply the new rules
  destinationSheet.setConditionalFormatRules([ordersHighlightRule, transfersHighlightRule]);

  // Clear the contents of the "Complete" sheet after row 1
  sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).clearFormat();
  sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).clearContent();

  // Reset the filter for the "Complete" sheet
  sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getFilter().remove();
  sourceSheet.getRange(1, 1, 1, columnHeaders.length).createFilter(); 
}
