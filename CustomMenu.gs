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

      // Move the values to the Transfers sheet with additional information
      Logger.log("Appending row to Transfers Sheet");
      transfersSheet.appendRow(['', ...data[i].slice(0, 9), formattedDate, quantity, employee, '', data[i][orderingNotesColumnP - 1]]);
      
      // Add checkboxes for the Undo and Complete columns
      transfersSheet.getRange(transfersSheet.getLastRow(), 1).setDataValidation(checkBox);
      transfersSheet.getRange(transfersSheet.getLastRow(), 14).setDataValidation(checkBox);
      
      // Delete the row from the source sheet
      orderSheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1 
    };
  };

  // Add a thick black border to the bottom of the final row
  transfersSheet.getRange(transfersSheet.getLastRow(), 1, 1, transfersSheet.getLastColumn()).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

};

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
        completeSheet.appendRow(['', ...data[i].slice(0, 9), formattedDate, quantity, employee, '', data[i][orderingNotesColumnP - 1], "O"]);

        // Add checkboxes for the Undo column
        completeSheet.getRange(completeSheet.getLastRow(), 1).setDataValidation(checkBox);

        // Set bg color based on which sheet the row came from
        completeSheet.getRange(completeSheet.getLastRow(), 1, 1, completeSheet.getLastColumn()).setBackground(completeColor);

        // Delete the row from the source sheet
        orderSheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1 
      }
    }   

    // Add a thick black border to the bottom of the final row
    completeSheet.getRange(completeSheet.getLastRow(), 1, 1, completeSheet.getLastColumn()).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  } else if (sName == "Transfers") {
    let transferQuantityColumnL = 12; // Column L (Quantity from Transfers Sheet)
    let transferSendingEmpColumnM = 13; // Column M (Sending Employee from Transfers Sheet)
    let transferCheckboxColumnN = 14; // Column N (Checkbox from Transfers Sheet)
    let transferNotesColumnO = 15; // Column O (Notes from Transfers Sheet)
    let dataRange = transfersSheet.getRange(2, 1, transfersSheet.getLastRow() - 1, transfersSheet.getLastColumn());
    let data = dataRange.getValues();
    transfersSheet.appendRow(['']);
    // Prompt the user for Employee Initials
    let employee = Browser.inputBox('Enter Employee Initials');

    // Iterate through each row
    for (let i = data.length - 1; i >= 0; i--) {
      let transferCheckboxValueM = data[i][transferCheckboxColumnN - 1]; // Adjust the index for JavaScript's 0-based indexing

      // Check if the checkbox is checked in column M
      if (transferCheckboxValueM === true || transferCheckboxValueM === 'TRUE') {
        
        // Move the values to the "Complete" sheet with additional information
        completeSheet.appendRow(['', ...data[i].slice(1, 10), formattedDate, data[i][transferQuantityColumnL - 1], employee, data[i][transferSendingEmpColumnM - 1], data[i][transferNotesColumnO - 1], "T"]);

        // Add checkboxes for the Undo column
        let newRow = completeSheet.getLastRow();
        completeSheet.getRange(newRow, 1).setDataValidation(checkBox);

        // Set bg color based on which sheet the row came from
        completeSheet.getRange(completeSheet.getLastRow(), 1, 1, completeSheet.getLastColumn()).setBackground(transferColor);

        // Delete the row from the source sheet
        transfersSheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1 
      };
    };
    transfersSheet.deleteRow(transfersSheet.getLastRow());
    
    // Add a thick black border to the bottom of the final row
    transfersSheet.getRange(transfersSheet.getLastRow(), 1, 1, transfersSheet.getLastColumn()).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  } else { 
    console.log("Sheet name: " + sName)
    return; 
  };

};

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
};

function moveUndoRows() {
  let sheet = ss.getActiveSheet();
  let dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  let data = dataRange.getValues();
  let undoCheckboxColumnA = 1;
  let notesColumnO = 15;
  let orderOrTransferColumnP = 16; // Column P corresponds to index 15 in the array
  
  // Check if we are running on the "Complete" sheet
  if (sheet.getName() === 'Complete') {
    // Iterate through each row
    for (let i = data.length - 1; i >= 0; i--) {
      let undoCheckboxValueA = data[i][undoCheckboxColumnA - 1];
      let orderOrTransferValue = data[i][orderOrTransferColumnP - 1]; // Value in column P ("O" or "T")

      // Check if the checkbox is checked in column A
      if (undoCheckboxValueA === true || undoCheckboxValueA === 'TRUE') {
        if (orderOrTransferValue === 'O') {
          // Move the values to the Ordering List sheet
          orderSheet.appendRow([...data[i].slice(1, 10), '', '', '', '', '', '', data[i][notesColumnO - 1]]);

          // Remove data validation for the conditional drop downs
          orderSheet.getRange(2, 3, orderSheet.getLastRow() - 1, 2).clearDataValidations();

          // Add checkboxes for the New, Nomo, and Req columns
          orderSheet.getRange(orderSheet.getLastRow(), 10).setDataValidation(checkBox); // Column J - New
          orderSheet.getRange(orderSheet.getLastRow(), 13).setDataValidation(checkBox); // Column M - Nomo
          orderSheet.getRange(orderSheet.getLastRow(), 14).setDataValidation(checkBox); // Column N - Req

        } else if (orderOrTransferValue === 'T') {
          // Move the values to the Transfers sheet
          transfersSheet.appendRow(['', ...data[i].slice(1, 12), data[i][13], "", data[i][notesColumnO - 1]]);

          // Add checkboxes for Undo and Complete columns
          transfersSheet.getRange(transfersSheet.getLastRow(), 1).setDataValidation(checkBox); // Column A - Undo
          transfersSheet.getRange(transfersSheet.getLastRow(), 14).setDataValidation(checkBox); // Column N - Complete
        };
                // Delete the row from the original sheet
        sheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1
      };
    };
    // Add a thick black border to the bottom of the final row
    transfersSheet.getRange(transfersSheet.getLastRow(), 1, 1, transfersSheet.getLastColumn()).setBorder(null, null, true, null, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  } else if (sheet.getName() === 'Transfers') {
    // Iterate through each row
    for (let i = data.length - 1; i >= 0; i--) {
      let undoCheckboxValueA = data[i][undoCheckboxColumnA - 1];
      
      // Check if the checkbox is checked in column A
      if (undoCheckboxValueA === true || undoCheckboxValueA === 'TRUE') {
        
        // Move the values to the Ordering List sheet
        orderSheet.appendRow([...data[i].slice(1, 10), '', '', '', '', '', '', data[i][notesColumnO - 1]]);

        // Remove data validation for the conditional drop downs
        orderSheet.getRange(2, 3, orderSheet.getLastRow() - 1, 2).clearDataValidations();

        // Add checkboxes for the New, Nomo, and Req columns
        orderSheet.getRange(orderSheet.getLastRow(), 10).setDataValidation(checkBox); // Column J - New
        orderSheet.getRange(orderSheet.getLastRow(), 13).setDataValidation(checkBox); // Column M - Nomo
        orderSheet.getRange(orderSheet.getLastRow(), 14).setDataValidation(checkBox); // Column N - Req

        // Delete the row from the original sheet
        sheet.deleteRow(i + 2); // Adding 2 because the loop starts from index 0 and row numbering starts from 1
      }
    };
  };
};

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
      range.sort([{column: 4, ascending: true}, {column: 5, ascending: true}, {column: 6, ascending: true}, {column: 7, ascending: true}]);
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
};

// Creates the custom menu option to run the custom functions
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Alphabetical Sort', 'alphabeticalSort')
    .addItem('Move Transfers', 'moveTransferRows')
    //.addItem('Move MF for NP Vending Transfers', 'moveVendingTransfers') | WORK IN PROGRESS
    .addItem('Move Nomo Rows', 'moveNomoRows')
    .addItem('Move Completed Rows', 'moveCompleteRows')
    .addItem('Undo Moved Rows', 'moveUndoRows')
    .addItem('Show Older Ordering List Dates', 'oldOrderingEntries')
    .addToUi();
  oldOrderingEntries();
};
