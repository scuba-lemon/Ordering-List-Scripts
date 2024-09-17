/*
-------------------------------------------------
Functions Index
-----

On Edit Functions:
 - onEditTriggerDropDown(e): Generates drop-down menus and inserts checkboxes if the cell containing the category is edited on the Ordering List sheet
 - onEditTriggerDevices(e): Generates drop-down menus and inserts checkboxes if the cell containing the category is edited on the Ordering List sheet
 - onEditUpdateDropdowns(e): Updates the drop down menus for all categories and subcategories if the column containing the list is edited on the Drop Downs sheet
 - onEditUpdateDevices(e): Updates the drop down menus for devices if the column containing the list is edited on the Devices sheet
 - onEditFormatRules(e): Implement formatting rules on the Ordering List sheet based on the column that is edited
 - insertListDate(e): For the Ordering List Sheet - Automatically inserts today's date into the list date cell whenever a new row is added
 - insertNPListDate(e): For the MF for NP Sheet - Automatically inserts today's date into the list date cell whenever a liquid is added or strikes through the row if the complete date is filled, and removes values and formatting if the cell is empty
 - onEditVendingFormat(e): Insert checkboxes on the MF for NP Vending Sheet when a new row is added
 - onEditSheetCheck(e): Runs the correct functions with a 10 second time-out, based on which sheet is edited 
 - onEdit(e): General function to run all On Edit functions with a time-out of 10 seconds

Helper Functions:
- getLastDataRowInColumn(columnName, sheet): Finds the last row containing values within a column
- timeoutFunction(func, timeout, functionName): Executes a passed function, with a timeout set with miliseconds - kills the script if it takes too long to execute
- makeArray(columnLetter, columnNumber, sheet): Gets all the values from the column, makes an array, and alphabetizes the array
- createCheckBoxes(currentRow): Generates checkboxes for the New, Nomo, and Req columns, so that the sheet can use a filter without including the empty rows that would only contain checkboxes
- findBrands(category): Assigns the correct brands for each category
- createBrandsMenu(): Generates a drop down menu for brands based on the category selected
- findDevices(brand): Assigns the correct devices for each brand
- createDevicesMenu(): Generates a drop down menu for devices based on the brand
- findSubCategory(category): Assigns the correct subcategory for each category
- createSubCategoryMenu(): Generates a drop down menu for subcategories based on the category selected
- transposeData(): Function to transpose five columns of the Ordering List into an object, excluding empty rows and converting values to lowercase
- transposeNomos(): Function to transpose five columns of the spreadsheet into an object, excluding empty rows and converting values to lowercase
- checkForDuplicateRow(e, products): Checks for duplicate listings when a new item is added - if a duplicate is found it returns the row index of the original and duplicate rows
- highlightDuplicateRows(originalRowIndex, duplicateRowIndex): Takes the indices returned by the checkForDuplicateRow function, and highlights the row based on the index returned for duplicate rows on the Order Sheet
- highlightDupicateNomos(originalRowIndex, duplicateRowIndex): Takes the indices returned by the checkForDuplicateRow function, and highlights the row based on the index returned for duplicate rows from the Nomo Sheet
- compareDates(valueInColumnI, row): Compares the order date to today's date, and if the date is older than 14 days, the cells will be highlighted
- oldOrderingEntries(): Runs on open to check for old ordering dates
- parseDate(dateStr): Function to parse date strings in various formats
- checkCheckboxesInRow(row): Function to check if any checkboxes in columns J, M, and N are checked for a given row and return the column number of the first checked checkbox
- highlightCheckForBoxes(checkedColumn, row): Function to highlight based on checked column

Custom Menu:
- moveTransferRows(): Moves rows to the Transfers page
- moveCompleteRows(): Moves rows to the Complete page
- moveNomoRows(): Moves rows with checked boxes to the Nomo page
- moveUndoRows(): Moves rows with checked boxes to the previous sheet
- alphabeticalSort(): Sorts the data alphabetically
- onOpen(): Creates the custom menu option to run the custom functions

Requests and Rejects:
- transposeRejects(): Function to transpose five columns of the spreadsheet into an object, excluding empty rows and converting values to lowercase
- formatData(e): Function which gets the data whenever a new request is submitted via the form, then formats and cleans up the data, and prepares it to be appended to the Requests sheet
- appendToRequestSheet(data): Appends a new row to the Requests sheet whenever an item is submitted via the form
- formatNewRequest(data): Finishes the formatting on the Requests sheet after an item is added
- onFormSubmit(e): Executes certain functions whenever the "Requests" Form is submitted
- onCheckMoveRow(e): Moves the row to either the Ordering List or Rejected Requests sheet if a checkbox is marked in column A or B

Archiving:
- moveAndClearCompleteSheet(): Automatically moves all of the information on the Complete sheet to an archive - runs on the 1st of each month at midnight

-------------------------------------------------
Ref
  1. A
  2. B
  3. C
  4. D
  5. E
  6. F
  7. G
  8. H
  9. I
  10. J
  11. K
  12. L
  13. M
  14. N
  15. O
  16. P
  17. Q
  18. R
  19. S
  20. T
  21. U
  22. V
  23. W
  24. X
  25. Y
  26. Z
*/
