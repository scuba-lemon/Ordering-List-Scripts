/*
-------------------------------------------------
Functions Index
-----
Helper Functions:
---
  - getLastDataRowInColumn(columnName, sheet)
    - General helper - retrieves the index for the last row which contains a value within a column
      
  - timeoutFunction(func, timeout, functionName)
    - General helper - error handling to prevent accidental infinite loops preventing a function from terminating
      
  - makeArray(columnLetter, columnNumber, sheet)
    - General helper - transposes spreadsheet data into a javascript array
      
  - createCheckBoxes(currentRow)
    - General helper and formatting rule - inserts a checkbox
      
  - findBrands(category)
    - Helper for createBrandsMenu - logic to retrieve the drop down menu values for brands after a category has been selected from the categories drop down options
      
  - createBrandsMenu()
    - Formatting rule - inserts a drop down menu with the values retrieved by findBrands
      
  - findDevices(brand)
    - Helper for createDevicesMenu - logic to retrieve the drop down menu values for device manufacturers after a brand has been selected from the brands drop down options - used when the category is Devices
      
  - createDevicesMenu()
    - Formatting rule - inserts a drop down menu with the values retrieved by findDevices
      
  - findSubCategory(category)
    - Helper for createSubCategoryMenu - logic to retrieve the drop down menu values for subcategories after a category has been selected from the categories drop down options
      
  - createSubCategoryMenu()
    - Formatting rule - inserts a drop down menu with the values retrieved by findSubCategory
      
  - transposeData()
    - Helper for checkForDuplicateRow - gets values and data from the sheet for comparison when checking for duplicates
      
  - transposeNomos()
    - Helper for checkForDuplicateRow - gets data to compare when checking for duplicates
      
  - checkForDuplicateRow(e, products)
    - Helper for highlightDuplicateRows and highlightDupicateNomos - logic to check for duplicate rows
      
  - highlightDuplicateRows(originalRowIndex, duplicateRowIndex)
    - Formatting rule - highlights duplicates when a new row is added
      
  - highlightDupicateNomos(originalRowIndex, duplicateRowIndex)
    - Formatting rule - highlights duplicates when a new row is added
      
  - compareDates(valueInColumnI, row)
    - Helper for oldOrderingEntries - gets data to compare when checking for old ordering dates
      
  - oldOrderingEntries()
    - Formatting rule - checks for and highlights dates which are more than 20 days old
      
  - parseDate(dateStr)
    - Helper for oldOrderingEntries - converts the ordering date value into a string in order to compares to today's date
      
  - checkCheckboxesInRow(row)
    - Helper for highlightCheckForBoxes - finds rows with checked boxes
      
  - highlightCheckForBoxes(checkedColumn)
    - Formatting rule - highlights rows that have checked boxes for visual color-coding

-----
On Edit Functions:
---
  - onEditVendingCheckbox(e)
    - Executor - inserts a checkbox when a new row is added for vending

  - onEditTriggerDropDown(e)
    - Executor - inserts a drop down with options for the brand, based on the category in the previous drop down

  - onEditTriggerDevices(e)
    - Executor - inserts a drop down with options for the device, based on the category and brand in the previous drop downs

  - onEditUpdateDropdowns(e)
    - Executor - updates the options within the drop down menus for brands when any of the lists with the values for the menus are edited

  - onEditUpdateDevices(e)
    - Executor - updates the options within the drop down menus for devices when any of the lists with the values for the menus are edited

  - onEditFormatRules(e)
    - Logic - applies formatting rules when the correct conditions are met

  - insertListDate(e)
    - Executor - Inserts today's date when a new row is added to the Ordering List sheet

  - insertNPListDate(e)
    - Executor - Inserts today's date when a new row is added to the MF for NP sheet

  - onEditVendingFormat(e)
    - Logic - applies formatting rules when the correct conditions are met

  - onEditSheetCheck(e)
    - Logic - applies formatting rules based on the sheet being edited

  - onEdit(e)
    - Executor - runs all onEdit subfunctions with a time-out for error handling

-----
Custom Menu:
---
  - moveTransferRows()
    - Tool - moves rows where the transfer quantity cell has a value from the Ordering List sheet to the Transfers sheet
    
  - moveCompleteRows()
    - Tool - moves rows where the complete quantity cell has a value from the Ordering List or Transfer sheet to the Complete sheet
      
  - moveNomoRows()
    - Tool - move rows where the checkbox for in the column for "Nomo" has been checked from the Orderling List sheet to the Nomo sheet
    
  - alphabeticalSort()
    - Tool - resets the filters on the Ordering List sheet and then sorts the data alphabetically
    
  - onOpen()
    - Executor - on open, creates a menu with the above tools as options to run from the front-end, and runs alphabeticalSort

-----
Requests: 
---
  - transposeRejects()
    - Helper for formatData - transposes data and values from the Requests and Rejects sheets for comparison
    
  - formatData(e)
    - Formatting rule - gets data from submissions received by the Requests form, reformats and prepares the data to match the spreadsheet structure, then checks to see if the new submission matches any entries on the Rejects sheet and highlights the row if the new submission matches any of the rejected requests
    
  - appendToRequestSheet(data)
    - Executor - appends new entries to the Requests sheet when the Requests form receives a new submission
    
  - formatNewRequest(data)
    - Formatting rule - finishes the formatting on the Requests sheet after a new entry is added
    
  - onFormSubmit(e)
    - Executor - receives data from the Requests form when the form is submitted
    
  - onCheckMoveRow(e)
    - Executor - moves rows from the Requests sheet to either the Ordering List or Rejects sheet when the appropriate checkbox is selected
    
-----
Archiving:
---
  - moveAndClearCompleteSheet()
    - Executor - copies the data from the Complete sheet to an archive at the end of each month, then clears the complete sheet

*/
