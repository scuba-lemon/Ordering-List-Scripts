// Moves rows to the Transfers page
function moveTransferRows() {
  let dataRange = orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, orderSheet.getLastColumn());
  let data = dataRange.getValues();
  let orderingQuantityColumnL = 12; // Column L (Cell from Ordering Sheet)
  let orderingNotesColumnP = 16; // Column P (Notes from Ordering Sheet)

  // Prompt the user for Employee Initials
  let employee = Browser.inputBox('Enter Employee Initials');

  // Iterate through each row in reverse order
  for (let i = data.length - 1; i >= 0; i--) {
    let orderingQuantityValueL = data[i][orderingQuantityColumnL - 1]; // Adjust the index for JavaScript's 0-based indexing

    // Check if the cell contains a value in column L
    if (orderingQuantityValueL !== "") {
      Logger.log("Found transfer on row #" + i);
      // Save the quantity for the item
      let quantity = orderingQuantityValueL;

      // Move the values to the "Transfers" sheet with additional information
      Logger.log("Appending row to Transfers Sheet");
      transfersSheet.appendRow([...data[i].slice(0, 8), formattedDate, quantity, employee, '', data[i][orderingNotesColumnP - 1]]);
      
      // Select the cell in column L of the current row on the "Transfers" sheet and add a checkbox for the "Received" Column
      let transferredRow = transfersSheet.getLastRow();
      let checkboxCell = transfersSheet.getRange(transferredRow, 12);
      checkboxCell.setDataValidation(checkBox);
      
      // Delete the row from the source sheet
      orderSheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1 
    }
  }

  // Add a thick black border to the bottom of the final row
  let lastRow = transfersSheet.getLastRow();
  let lastColumn = transfersSheet.getLastColumn();
  let lastRowRange = transfersSheet.getRange(lastRow, 1, 1, lastColumn);
  lastRowRange.setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

}

// Moves rows to the Complete page
function moveCompleteRows() {
  let sName = activeSheet.getSheetName();
  
  if (sName == "Ordering List") {
    let orderingNotesColumnP = 16; // Column P (Notes from Ordering Sheet)
    let orderingQuantityColumnO = 15; // Column O (Cell from Ordering Sheet)
    let dataRange = orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, orderSheet.getLastColumn());
    let data = dataRange.getValues();

    // Prompt the user for Employee Initials
    let employee = Browser.inputBox('Enter Employee Initials');

    // Iterate through each row
    for (let i = data.length - 1; i >= 0; i--) {
      let orderingQuantityValueO = data[i][orderingQuantityColumnO - 1]; // Adjust the index for JavaScript's 0-based indexing

      // Check if the cell contains a value in column O
      if (orderingQuantityValueO !== "") {
        // Save the quantity for the item
        let quantity = orderingQuantityValueO;

        // Move the values to the "Complete" sheet with additional information
        completeSheet.appendRow([...data[i].slice(0, 9), formattedDate, quantity, employee, '', data[i][orderingNotesColumnP - 1], "O"]);
        completeSheet.getRange(completeSheet.getLastRow(), 1, 1, completeSheet.getLastColumn()).setBackground(completeColor);

        // Delete the row from the source sheet
        orderSheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1 
      }
    }   

    // Add a thick black border to the bottom of the final row
    let lastRow = completeSheet.getLastRow();
    let lastColumn = completeSheet.getLastColumn();
    let lastRowRange = completeSheet.getRange(lastRow, 1, 1, lastColumn);
    lastRowRange.setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  } else if (sName == "Transfers") {
    let transferQuantityColumnJ = 10; // Column J (Quantity from Transfers Sheet)
    let transferSendingEmpColumnK = 11; // Column K (Quantity from Transfers Sheet)
    let transferCheckboxColumnL = 12; // Column L (Checkbox from Transfers Sheet)
    let transferNotesColumnM = 13; // Column M (Notes from Transfers Sheet)
    let dataRange = transfersSheet.getRange(2, 1, transfersSheet.getLastRow() - 1, transfersSheet.getLastColumn());
    let data = dataRange.getValues();
    transfersSheet.appendRow(['']);
    // Prompt the user for Employee Initials
    let employee = Browser.inputBox('Enter Employee Initials');

    // Iterate through each row
    for (let i = data.length - 1; i >= 0; i--) {
      let transferCheckboxValueM = data[i][transferCheckboxColumnL - 1]; // Adjust the index for JavaScript's 0-based indexing

      // Check if the checkbox is checked in column M
      if (transferCheckboxValueM === true || transferCheckboxValueM === 'TRUE') {
        
        // Move the values to the "Complete" sheet with additional information
        completeSheet.appendRow([...data[i].slice(0, 9), formattedDate, data[i][transferQuantityColumnJ - 1], employee, data[i][transferSendingEmpColumnK - 1], data[i][transferNotesColumnM - 1], "T"]);
        completeSheet.getRange(completeSheet.getLastRow(), 1, 1, completeSheet.getLastColumn()).setBackground(transferColor);

        // Delete the row from the source sheet
        transfersSheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1 
      }
    }
    transfersSheet.deleteRow(transfersSheet.getLastRow());
    
    // Add a thick black border to the bottom of the final row
    let lastRow = completeSheet.getLastRow();
    let lastColumn = completeSheet.getLastColumn();
    let lastRowRange = completeSheet.getRange(lastRow, 1, 1, lastColumn);
    lastRowRange.setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  } else { 
    console.log("Sheet name: " + sName)
    return; 
  };

}

// Moves rows with checked boxes to the Nomo page
function moveNomoRows() {
  let dataRange = orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, orderSheet.getLastColumn());
  let data = dataRange.getValues();
  let orderingCheckboxColumnM = 13; // Column M (Checkbox from Ordering Sheet)
  
  // Iterate through each row
  for (let i = data.length - 1; i >= 0; i--) {
    let orderingCheckboxValueM = data[i][orderingCheckboxColumnM - 1]; // Adjust the index for JavaScript's 0-based indexing

    // Check if the checkbox is checked in column M
    if (orderingCheckboxValueM === true || orderingCheckboxValueM === 'TRUE') {
      // Move the values in columns A-E to the "Transfers" sheet
      nomoSheet.appendRow([...data[i].slice(0, 5), data[i][data[i].length - 1]]);
    }

    // If checkbox is checked, delete the row from the source sheet
    if (orderingCheckboxValueM === true || orderingCheckboxValueM === 'TRUE') {
      orderSheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1
    }
  }
} 

// Sorts the data alphabetically
function alphabeticalSort() {
  let sName = activeSheet.getSheetName();
  let range;

  if (sName == "Notes" || sName == "MF for NP") {
    Logger.log("Alphabetizing attempt on invalid sheet.");
    return;
  } else {
    // Check if there are any existing filters on the sheet & add a filter to the sheet
    if (activeSheet.getFilter()) { 
        activeSheet.getFilter().remove(); 
        activeSheet.getDataRange().createFilter();
      } else {
        activeSheet.getDataRange().createFilter();
      }
    
    // Define the range based on sheet name
    if (sName == "Ordering List") {
      Logger.log("Alphabetizing the Ordering List Sheet");
      let lastRow = getLastDataRowInColumn("A", activeSheet);
      let numColumns = orderSheet.getLastColumn();
      let emptyRow = Array(numColumns).fill('');
      orderSheet.appendRow(emptyRow);
      range = orderSheet.getRange(2, 1, lastRow, orderSheet.getLastColumn());
      range.sort([{column: 2, ascending: true}, {column: 3, ascending: true}, {column: 4, ascending: true}, {column: 5, ascending: true}]);
    } else if (sName == "Transfers" || sName == "Complete" || sName == 'Nomo') {
      let lastRow = getLastDataRowInColumn("A", activeSheet);
      Logger.log("Alphabetizing the Transfers, Complete, or Nomo Sheet");
      range = activeSheet.getRange(2, 1, lastRow, activeSheet.getLastColumn());
      range.sort([{column: 3, ascending: true}, {column: 4, ascending: true}, {column: 5, ascending: true}, {column: 6, ascending: true}]);
    } else if (sName == "Drop Downs") {
      let lastRow = getLastDataRowInColumn("D", activeSheet);
      Logger.log("Alphabetizing the Drop Downs Sheet");
      range = activeSheet.getRange(2, 1, lastRow, activeSheet.getLastColumn());
      for (let col = 1; col <= activeSheet.getLastColumn(); col++) {
        range.sort({column: col, ascending: true});
      }
    } else if (sName == "Devices") {
      let lastRow = getLastDataRowInColumn("M", activeSheet);
      Logger.log("Alphabetizing the Devices Sheet");
      range = activeSheet.getRange(2, 1, lastRow, activeSheet.getLastColumn());
      for (let col = 1; col <= activeSheet.getLastColumn(); col++) {
        range.sort({column: col, ascending: true});
      }
    } else if (sName == "Requests") {
      let lastRow = getLastDataRowInColumn("A", activeSheet);
      Logger.log("Alphabetizing the Requests Sheet");
      range = activeSheet.getRange(2, 1, lastRow, activeSheet.getLastColumn());
      for (let col = 1; col <= activeSheet.getLastColumn(); col++) {
        range.sort([{column: 5, ascending: true}, {column: 6, ascending: true}, {column: 7, ascending: true}, {column: 8, ascending: true}]);
      }
    } else if (sName == "Rejected Requests") {
      let lastRow = getLastDataRowInColumn("A", activeSheet);
      Logger.log("Alphabetizing the Rejects Sheet");
      range = activeSheet.getRange(2, 1, lastRow, activeSheet.getLastColumn());
      for (let col = 1; col <= activeSheet.getLastColumn(); col++) {
        range.sort([{column: 2, ascending: true}, {column: 3, ascending: true}, {column: 4, ascending: true}, {column: 5, ascending: true}]);
      }
    }
  }
}

// Creates the custom menu option to run the custom functions
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Alphabetical Sort', 'alphabeticalSort')
    .addItem('Move Transfers', 'moveTransferRows')
    //.addItem('Move MF for NP Vending Transfers', 'moveVendingTransfers') | WORK IN PROGRESS
    .addItem('Move Nomo Rows', 'moveNomoRows')
    .addItem('Move Completed Rows', 'moveCompleteRows')
    .addItem('Show Older Ordering List Dates', 'oldOrderingEntries')
    .addToUi();
  oldOrderingEntries();
}
