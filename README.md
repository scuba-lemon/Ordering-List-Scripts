/*
-------------------------------------------------
Function File Locations
-----

Helper Functions:
  - getLastDataRowInColumn(columnName, sheet)
  - timeoutFunction(func, timeout, functionName)
  - makeArray(columnLetter, columnNumber, sheet)
  - createCheckBoxes(currentRow)
  - findBrands(category)
  - createBrandsMenu()
  - findDevices(brand)
  - createDevicesMenu()
  - findSubCategory(category)
  - createSubCategoryMenu()
  - transposeData()
  - transposeNomos()
  - checkForDuplicateRow(e, products)
  - highlightDuplicateRows(originalRowIndex, duplicateRowIndex)
  - highlightDupicateNomos(originalRowIndex, duplicateRowIndex)
  - insertListDate(e)
  - insertNPListDate(e)
  - compareDates(valueInColumnI, row)
  - oldOrderingEntries()
  - parseDate(dateStr)
  - checkCheckboxesInRow(row)
  - highlightCheckForBoxes(checkedColumn)

On Edit Functions:
  - onEditVendingCheckbox(e)
  - onEditTriggerDropDown(e)
  - onEditTriggerDevices(e)
  - onEditUpdateDropdowns(e)
  - onEditUpdateDevices(e)
  - onEditFormatRules(e)
  - onEditSheetCheck(e)
  - onEdit(e)

Custom Menu:
  - moveTransferRows()
  - moveCompleteRows()
  - moveNomoRows()
  - alphabeticalSort()
  - onOpen()

Requests: 
  - transposeRejects()
  - formatData(e)
  - appendToRequestSheet(data)
  - formatNewRequest(data)
  - onFormSubmit(e)
  - onCheckMoveRow(e)

Archiving:
  - moveAndClearCompleteSheet()

-------------------------------------------------
Changelog
-----
3/18/24
- Created Changelog, Bugs, & Features Doc.

-----
4/29/24
- Began working on new Requests form/sheet prototype system.

-----
5/2/24
- Updated the CompleteSheetArchiving code to update the headers - previously the code was missing headers for the Index, Employee (sending), and T/O columns, so the archive's header labels were offset by 1 column, and not correctly aligned with the values in the columns. Future Archive sheets should have the correct headers automatically generated. I have corrected the headers for the archives of completed orders from March and May of 2024.

-----
5/4/24
- Found that somehow the Joyetech brand name and devices had been removed from the drop down menu options that automatically populate the fields. Don't know if someone did that on purpose, but I fixed it, and put it back. >:|

-----
5/3/24
- Continued working on new Requests form/sheet system.

-----
5/5/24
- Created "Rejected Requests" sheet and developed the automatic row moving/management functions for the Requests sheet.

-----
5/6/24

- Continued working on the Requests sheet - resolved duplicate function execution issue, updated formatting on the Ordering List after moving the row.
- Discovered that onEdit() will run whether or not the trigger is set up in the triggers panel of Apps Scripts, hence my duplicate execution issue. Removed the trigger from Apps Script's triggers panel, keeping the in-script function only.

-----
5/9/24
- Fixed the transposeRejects skipping the 2nd row of the spreadsheet so that if a rejected item is added to the requests sheet, it will be highlighted with the "nomo" colors.
- Finished the combined notes transposing when a row is moved.
- Finished working on the new Requests/Rejects sheets.

-----
5/10/24
- Updated the createCheckboxes() function to no longer rely on the .getCurrentCell() method, so that the function is re-usable within the requests code.
- Broke the Scripts file into seperate files for Requests, On Edit Functions, Helper Functions, Custom Menu Functions, Variables, and Archiving bc it is getting a little out of hand trying to keep everything in one Scripts.gs file.
- Created the today and formattedDate variables to automatically populate the date and improve code reusability; used in the moveTransferRows, moveCompleteRows, insertListDate (used within onEditTriggerDropDown), and onCheckMoveRow.

-----
5/11/24
- Updated the insertListDate, createCheckboxes, onEditTriggerDropDown, and onCheckMoveRow functions so that the automatic date insertion and checkboxes will be added automatically to the order sheet, regardless of whether you are working on the last row or not. 

-----
5/20/24
- Removed the on open trigger, since it works the same way as the on edit triggers - automatically works based on the name via native Google Scripts functionality. 
- Added two new functions oldOrderingEntries and parseDate, and added an instance of oldOrderingEntries to the onOpen function, so that entries with ordering dates older than 2 weeks will automatically get a thick black border around the list & ordering dates. Also added a custom menu option to run the function, so the sheet doesn't HAVE to be reopened to run it outside of scripts. 

-----
5/25/24
- Officially implemented the new requests form & sheets system! 
- Edited the code which looks for old ordering dates, so that it will remove borders if the date is erased or updated to within 2 weeks of the current date. Also added the compareDates function in order to improve code reusablility. 

-----
6/11/24
- Changed the function which borders old ordering dates to highlight instead. 

-----
6/12/24
- Added the insertNPListDate function to automatically populate the list date when liquid is added to MF for NP.

-----
6/15/24
- Modified the onEditFormatRules to include the new helper functions to correct the old order date highlighting rules.
- Added the checkCheckboxesInRow and highlightCheckForBoxes functions to correct the old order date highlighting rules.

-----
6/19/24
- Removed all index number logic and code from the custom menu and archiving functions, so that the transfers, complete, nomo, and archive won't include an index number column. 

-----
6/28/24
- Modified the oldOrderingEntries function to end the function when the loop reaches the last row. 

-----
7/29/24
- Updated the alphabeticalSort and onCheckMoveRow (for reqs/rejs) functions to resolve a sorting error by appending an empty row before sorting

*/
