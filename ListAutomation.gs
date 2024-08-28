function updateProductLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName("Sales");

  const data = salesSheet.getDataRange().getValues();
  const salesData = data.slice(1).map(row => ({
    date: row[0],        // Column A
    category: row[3],    // Column D
    itemName: row[4],    // Column E
    quantity: row[5],    // Column F
    variation: row[6],   // Column G
    modifier: row[8],    // Column I
    location: row[19]    // Column T
  }));

  createDateNamedSheet();
  
  // Venice
  createOrReplaceSheet(ss, "Venice");
  processList(salesData, "Mighty Fine Flavors", "Venice", "variation", true);
  formatList(ss.getSheetByName("Venice"));

  // Venice Liquid
  createOrReplaceSheet(ss, "Venice Liquid");
  processLiquidSalesData(salesData, ss.getSheetByName("Venice Liquid"));
  formatLiquidList(ss.getSheetByName("Venice Liquid"));

  // North Port
  createOrReplaceSheet(ss, "North Port");
  processList(salesData, "Mighty Fine Vape & Smoke | North Port", "North Port", "variation", true);
  formatList(ss.getSheetByName("North Port"));
};

function createDateNamedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formatDate = (date) => `${String(date.getMonth() + 1).padStart(2, '0')}.${String(date.getDate()).padStart(2, '0')}`;

  const recentDate = formatDate(findMostRecentDate());
  const oldestDate = formatDate(findOldestDate());
  const sheetName = `Date: ${oldestDate} - ${recentDate}`;

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith("Date")) ss.deleteSheet(sheet);
  });
  ss.insertSheet(sheetName);
};

function findMostRecentDate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
  const dates = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
  const recentDate = new Date(Math.max(...dates.filter(Boolean).map(date => new Date(date))));
  return recentDate;
};

function findOldestDate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
  const dates = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
  const oldestDate = new Date(Math.min(...dates.filter(Boolean).map(date => new Date(date))));
  return oldestDate;
};

function createOrReplaceSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  ss.insertSheet(sheetName);
};

function processList(salesData, filterValue, sheetName, sumBy, sortByCategory, filterBy = "location") {
  let filteredData = salesData.filter(item => 
    item[filterBy] === filterValue && 
    item.category !== "Mighty Fine" && 
    item.category !== "Mighty Fine NP"
  );
  let summedData = sumQuantities(filteredData, sumBy);
  if (sortByCategory) {
    summedData.sort((a, b) => a.category.localeCompare(b.category));
  }
  writeDataByLocation(summedData, sheetName);
};

function sumQuantities(data, key) {
  return Object.values(data.reduce((acc, item) => {
    if (!acc[item[key]]) {
      acc[item[key]] = {
        category: item.category,
        itemName: item.itemName,
        variation: item.variation,
        sum: 0
      };
    }
    acc[item[key]].sum += item.quantity;
    return acc;
  }, {}));
};

function writeDataByLocation(data, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  sheet.getRange("A1").activate().setValue("");
  let headers = [["Item", "Variation", "Category", "Qty"]];
  let rows = data.map(item => [item.itemName, item.variation, item.category, item.sum]);
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
};

function formatList(sheet) {
  let range = sheet.getDataRange();

  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment('center');
  sheet.getRange(1, 2, sheet.getLastRow(), 3).setHorizontalAlignment('center');
  sheet.setColumnWidth(4, 50);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(1, 250);
  
  range.createFilter();
  sheet.getRange('B1').activate().getFilter().sort(2, true);
  sheet.getRange('A1').activate().getFilter().sort(1, true);
  sheet.getRange('C1').activate().getFilter().sort(3, true);
  
  applyConditionalFormatting(sheet);
  hideUnsortedNomoItems(sheet);
};

function processLiquidSalesData(data, sheet) {
  const filteredData = data.filter(item => item.category === "Mighty Fine");
  const headers = [["Date", "Item", "Modifiers", "Qty"]];
  const rows = filteredData.map(item => [item.date, item.itemName, item.modifier, item.quantity]);
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
};

function formatLiquidList(sheet) {
  let range = sheet.getDataRange();

  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setHorizontalAlignment('center');
  sheet.setColumnWidth(4, 50);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(1, 100);

  range.createFilter();
  sheet.getRange('C1').activate().getFilter().sort(3, true);
  sheet.getRange('B1').activate().getFilter().sort(2, true);

  applyConditionalFormatting(sheet, true);
  hideUnsortedNomoItems(sheet);
  hideOldRows(sheet);
};

function applyConditionalFormatting(sheet, isLiquid = false) {
  const formatRange = sheet.getRange(isLiquid ? "C2:C" : "A2:B");
  const rules = sheet.getConditionalFormatRules();
  if (!isLiquid) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith("XL 3g - (").setBold(false).setBackground(null).setRanges([formatRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith("X").setBold(true).setBackground("#e2e2e2").setRanges([formatRange]).build()
    );
  } else {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextContains("30").setBold(true).setBackground("#e2e2e2").setRanges([formatRange]).build()
    );
  }
  sheet.setConditionalFormatRules(rules);
};

function hideUnsortedNomoItems(sheet) {
  const lastRow = sheet.getLastRow();
  const valuesA = sheet.getRange(`A2:A${lastRow}`).getValues();
  const valuesB = sheet.getRange(`B2:B${lastRow}`).getValues();
  
  valuesA.forEach((valueA, i) => {
    const valueB = valuesB[i][0].toString();
    if ((valueA[0] && valueA[0].toString().startsWith("00")) || 
        (valueB && valueB.startsWith("00"))) {
      sheet.hideRows(i + 2);
    } else if ((valueA[0].toString().startsWith("X") && !valueA[0].includes("XL 3g")) || 
               (valueB.startsWith("X") && !valueB.includes("XL 3g"))) {
      sheet.hideRows(i + 2);
    }
  });
};

function hideOldRows(sheet) {
  const recentDate = findMostRecentDate();
  const oneWeekAgo = new Date(recentDate);
  oneWeekAgo.setDate(recentDate.getDate() - 7);

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(`A2:A${lastRow}`).getValues();

  values.forEach((date, i) => {
    const rowDate = new Date(date[0]);
    if (rowDate < oneWeekAgo || rowDate > recentDate) {
      sheet.hideRows(i + 2);
    };
  });
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Custom Menu').addItem('Make the lists!', 'updateProductLists').addToUi();
};
