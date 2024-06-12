# ordering-list-scripts
Custom functions and code created to improve a spreadsheet that is used for tracking and managing inventory.
-----
Created by Makaela Wesner
-----

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
- Created Changelog, Bugs, & Features Doc 

-----
4/29/24
- Began working on new Requests form/sheet prototype system

-----
5/2/24
- Updated the CompleteSheetArchiving code to update the headers - previously the code was missing headers for the Index, Employee (sending), and T/O columns, so the archive's header labels were offset by 1 column, and not correctly aligned with the values in the columns. Future Archive sheets should have the correct headers automatically generated. I have corrected the headers for the archives of completed orders from March and May of 2024

-----
5/4/24
- Found that somehow the Joyetech brand name and devices had been removed from the drop down menu options that automatically populate the fields. Don't know if someone did that on purpose, but I fixed it, and put it back. >:|

-----
5/3/24
- Continued working on new Requests form/sheet system

-----
5/5/24
- Created "Rejected Requests" sheet and developed the automatic row moving/management functions for the Requests sheet

-----
5/6/24

- Continued working on the Requests sheet - resolved duplicate function execution issue, updated formatting on the Ordering List after moving the row
- Discovered that onEdit() will run whether or not the trigger is set up in the triggers panel of Apps Scripts, hence my duplicate execution issue. Removed the trigger from Apps Script's triggers panel, keeping the in-script function only

-----
5/9/24
- Fixed the transposeRejects skipping the 2nd row of the spreadsheet so that if a rejected item is added to the requests sheet, it will be highlighted with the "nomo" colors
- Finished the combined notes transposing when a row is moved
- Finished working on the new Requests/Rejects sheets

-----
5/10/24
- Updated the createCheckboxes() function to no longer rely on the .getCurrentCell() method, so that the function is re-usable within the requests code
- Broke the Scripts file into seperate files for Requests, On Edit Functions, Helper Functions, Custom Menu Functions, Variables, and Archiving bc it is getting a little out of hand trying to keep everything in one Scripts.gs file
- Created the today and formattedDate variables to automatically populate the date and improve code reusability; used in the moveTransferRows, moveCompleteRows, insertListDate (used within onEditTriggerDropDown), and onCheckMoveRow

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
- Added the insertNPListDate function to automatically populate the list date when liquid is added to MF for NP

-------------------------------------------------
Known Bugs
-----
3/14/24
- Alphabetical Sort function is not keeping rows intact when sorting - columns B-E look right, but the rows in the rest of the columns don't move. 
- Squashed 3/14/24 - Alphabetical sort function revised to use array within .sort() method on the ordering list, transfers, nomo, and complete sheets, instead of a for loop sorting each column individually. 
- The revised Alphabetical Sort function doesn't seem to want to alphabetize the drop downs sheet - still working on a fix, probably will need to change the for loop into an array like the other sheets.

-----
3/18/24
- The assignment of index numbers in column A on transfers, nomo, and complete sheets need to be turned into a function & corrected to ensure the numbers are sequential. Currently the newest entry is not necessarily going to be assigned the correct next index number. 

-----
5/9/24
- If too many rows are checked too quickly on the Requests sheet, the wrong rows may be moved. Hopefully should be caught by a try/catch statement, but there could still be bugs due to asynchronous functions - as far as I know that cannot really be avoided in Google Sheets Scripts, but I may do some more research on it in the future. In the meantime as long as the checkboxes are clicked at a reasonable speed it shouldn't be a problem.

-----
6/10/24
- The bordering for old ordering dates has some kind of issue with the alphabetical sort. No idea why, but cells are being bordered when they either aren't old enough to trigger the border, or the cell is empty. 

-----
6/12/24
- Tried modifying the old ordering dates function to highlight intead of border, and it didn't want to do anything when I was working on it. 

-------------------------------------------------
Planned/Suggested Features
-----
3/18/24
- Checkboxes on transfers, nomo, and complete sheets to send the row back to the ordering list, in case a row is added by mistake. 

-----
5/10/24
- Collin suggested some kind of "list count" for the requests sheet if we have an item requested more than once. For now, I feel like the nomo/reject highlighting accomplishes the same goal, but I'll think about how that could be implemented in the future. 
- Add the auto-date filling from the Requests transferring code to the code for moving Transfers and Complete rows, so there's one fewer fields to fill in. 

-------------------------------------------------
