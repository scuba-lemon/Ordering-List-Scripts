// Gets the active spreadsheet and all sheets within the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName('Ordering List');
  const mfForNpVendingSheet = ss.getSheetByName("MF for NP Vending");
  const mfForNpSheet = ss.getSheetByName("MF for NP");
  const requestSheet = ss.getSheetByName('Requests');
  const transfersSheet = ss.getSheetByName('Transfers');
  const completeSheet = ss.getSheetByName('Complete');
  const nomoSheet = ss.getSheetByName('Nomo');
  const dropDownSheet = ss.getSheetByName('Drop Downs');
  const deviceSheet = ss.getSheetByName('Devices');
  const rejectSheet = ss.getSheetByName('Rejected Requests');
  let activeSheet = ss.getActiveSheet();

// Initiates arrays and variables for later use
  let brands = ['N/A'];
  let devices = ['N/A'];
  let subCategory = ['N/A'];
  let currentCell = null;
  let checkBox = SpreadsheetApp.newDataValidation().requireCheckbox().build(); // Generates a checkbox
  let lock = false; // Variable to prevent concurrent function invocations
  let today = new Date();
  let formattedDate = today.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: '2-digit' });

// Highlight colors
  const transferColor = '#c9b8e6';
  const newColor = '#faf1a2';
  const nomoColor = '#db9595';
  const duplicateColor = '#b8cce6';
  const reqColor = '#eafde6'
  const completeColor = '#b1e0d3';
  const regularColor = '#bea6e0';
  const potentialRegularColor = '#e5d7fa';
  const oldColor = '#fad1a5';

// Initiates drop-down menu arrays - should only need to be run once, then the onEdit functions will keep the drop-down menus up to date.
  let deltaCos = makeArray('B', 2, dropDownSheet);
  let vapeCos = makeArray('C', 3, dropDownSheet);
  let eLiquid = makeArray('D', 4, dropDownSheet);
  let dispoCos = makeArray('E', 5, dropDownSheet);
  let accessoriesSubs = makeArray('F', 6, dropDownSheet);
  let deltaSubs = makeArray('G', 7, dropDownSheet);
  let stashSubs = makeArray('H', 8, dropDownSheet);

  let aspireDevs = makeArray('A', 1, deviceSheet);
  let dazzleafDevs = makeArray('B', 2, deviceSheet);
  let freemaxDevs = makeArray('C', 3, deviceSheet);
  let geekvapeDevs = makeArray('D', 4, deviceSheet);
  let horizontechDevs = makeArray('E', 5, deviceSheet);
  let joyetechDevs = makeArray('F', 6, deviceSheet);
  let juulDevs = makeArray('G', 7, deviceSheet);
  let kangertechDevs = makeArray('H', 8, deviceSheet);
  let lookahDevs = makeArray('I', 9, deviceSheet);
  let lostvapeDevs = makeArray('J', 10, deviceSheet);
  let motiXplayDevs = makeArray('K', 11, deviceSheet);
  let puffcoDevs = makeArray('L', 12, deviceSheet);
  let smokDevs = makeArray('M', 13, deviceSheet);
  let suorinDevs = makeArray('N', 14, deviceSheet);
  let uwellDevs = makeArray('O', 15, deviceSheet);
  let vaporessoDevs = makeArray('P', 16, deviceSheet);
  let voopooDevs = makeArray('Q', 17, deviceSheet);
  let yocanDevs = makeArray('R', 18, deviceSheet);

// Ref
  // 1. A
  // 2. B
  // 3. C
  // 4. D
  // 5. E
  // 6. F
  // 7. G
  // 8. H
  // 9. I
  // 10. J
  // 11. K
  // 12. L
  // 13. M
  // 14. N
  // 15. O
  // 16. P
  // 17. Q
  // 18. R
  // 19. S
  // 20. T
  // 21. U
  // 22. V
  // 23. W
  // 24. X
  // 25. Y
  // 26. Z
