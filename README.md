/*
-------------------------------------------------
Functions Index
-----
Helper Functions:
---
  - getLastDataRowInColumn(columnName, sheet)
    - General helper - retrieves the index for the last row in the column which contains a value
      
  - timeoutFunction(func, timeout, functionName)
    - General helper - error handling to prevent accidental infinite loops preventing a function from terminating
      
  - makeArray(columnLetter, columnNumber, sheet)
    - General helper for the functions which retrieve and update the values used to create drop down menus - transposes spreadsheet data into a javascript array
      
  - createCheckBoxes(currentRow)
    - Formatting rule to insert a checkbox
      
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
    - General helper - gets data to compare when checking for duplicates, helper for checkForDuplicateRow
      
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
    - Helper for oldOrderingEntries
      
  - checkCheckboxesInRow(row)
    - Helper for highlightCheckForBoxes - finds rows with checked boxes
      
  - highlightCheckForBoxes(checkedColumn)
    - Formatting rule - highlights rows that have checked boxes for visual color-coding

-----
On Edit Functions:
---
  - onEditVendingCheckbox(e)

  - onEditTriggerDropDown(e)

  - onEditTriggerDevices(e)

  - onEditUpdateDropdowns(e)

  - onEditUpdateDevices(e)

  - onEditFormatRules(e)

  - insertListDate(e)

  - insertNPListDate(e)

  - onEditVendingFormat(e)

  - onEditSheetCheck(e)

  - onEdit(e)

-----
Custom Menu:
---
  - moveTransferRows()
    
  - moveCompleteRows()
    
  - moveNomoRows()
    
  - alphabeticalSort()
    
  - onOpen()

-----
Requests: 
---
  - transposeRejects()
    
  - formatData(e)
    
  - appendToRequestSheet(data)
    
  - formatNewRequest(data)
    
  - onFormSubmit(e)
    
  - onCheckMoveRow(e)
    
-----
Archiving:
---
  - moveAndClearCompleteSheet()
    

*/
