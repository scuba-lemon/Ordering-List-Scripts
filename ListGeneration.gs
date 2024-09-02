function sumQuantities(data, key) {
  return Object.values(data.reduce((acc, item) => {
    const quantity = Number(item.quantity);
    if (isNaN(quantity)) {
      console.warn(`Skipping non-numeric quantity: ${item.quantity}`);
      return acc;
    }

    let sumKey;
    if (item.category === "Kits" || item.category === "Mech" || 
        item.category === "Mods" || item.category === "Tanks") {
      sumKey = item.itemName;
    } else {
      sumKey = key === "variation" 
        ? `${item.itemName} - ${item.variation}` 
        : item.itemName;
    }
    
    if (!acc[sumKey]) {
      acc[sumKey] = {
        category: item.category,
        itemName: item.itemName,
        variation: item.variation,
        sum: 0
      };
    }
    acc[sumKey].sum += quantity;
    return acc;
  }, {}));
};

function findDateRange() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sales');
  const dates = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
  const validDates = dates.filter(Boolean).map(date => new Date(date));
  
  if (validDates.length === 0) {
    return { recentDate: null, oldestDate: null };
  };

  const recentDate = new Date(Math.max(...validDates));
  const oldestDate = new Date(Math.min(...validDates));
  return { recentDate, oldestDate };
};

function createDateNamedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { recentDate, oldestDate } = findDateRange();
  if (!recentDate || !oldestDate) return;

  const formatDate = (date) => `${String(date.getMonth() + 1).padStart(2, '0')}.${String(date.getDate()).padStart(2, '0')}`;
  const sheetName = `Date: ${formatDate(oldestDate)} - ${formatDate(recentDate)}`;

  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith("Date")) ss.deleteSheet(sheet);
  });
  ss.insertSheet(sheetName);
};

function createOrReplaceSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

function writeDataByLocation(data, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  sheet.getRange("A1").activate().setValue("");
  let headers = [["Item", "Variation", "Category", "Qty"]];

  let rows = data.map(item => {
    let variation = String(item.variation);

    if (variation.toLowerCase() === "wed nov 20 2024 00:00:00 gmt-0500 (eastern standard time)") {
      variation = "11-20";
    };

    // Ensure that any numeric-looking variation names (like "11-20" or "0-10") are not interpreted as dates
    if (!isNaN(Date.parse(variation))) {
      variation = "'" + variation; // Prepend a single quote to force it to be interpreted as a string
    };

    if (!variation || variation === "Regular" || variation === "" || 
        item.category === "Kits" || item.category === "Mech" || 
        item.category === "Mods" || item.category === "Tanks") {
      let parts = item.itemName.split("- ");
      variation = parts.length > 1 ? parts[1] : "";
    };

    if (item.category === "Cones") {
      item.category = "SStash";
    };

    return [item.itemName, variation, item.category, item.sum];
  });

  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
};

function formatList(sheet) {
  sheet.getDataRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment('center');
  sheet.getRange(1, 2, sheet.getLastRow(), 3).setHorizontalAlignment('center');
  sheet.setColumnWidth(4, 50);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(1, 250);
  
  sheet.getDataRange().createFilter();
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
  sheet.getDataRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setHorizontalAlignment('center');
  sheet.setColumnWidth(4, 50);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(1, 100);

  sheet.getDataRange().createFilter();
  sheet.getRange('C1').activate().getFilter().sort(3, true);
  sheet.getRange('B1').activate().getFilter().sort(2, true);

  applyConditionalFormatting(sheet, true);
  hideUnsortedNomoItems(sheet);
  hideOldRows(sheet);
};

function applyConditionalFormatting(sheet, isLiquid = false) {
  const formatRange = sheet.getRange(isLiquid ? "B2:C" : "A2:B");
  const rules = sheet.getConditionalFormatRules();
  const newRules = [];

  if (!isLiquid) {
    const conditions = [
      { text: "XL 3g - (", bold: false, background: null },
      { text: "XROS", bold: false, background: null },
      { text: "Xros", bold: false, background: null },
      { text: "xros", bold: false, background: null },
      { text: "X", bold: true, background: "#e2e2e2" }
    ];
    conditions.forEach(cond => {
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextStartsWith(cond.text)
          .setBold(cond.bold)
          .setBackground(cond.background)
          .setRanges([formatRange])
          .build()
      );
    });
  } else {
    newRules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("30")
        .setBold(true)
        .setBackground("#e2e2e2")
        .setRanges([formatRange])
        .build()
    );
  };
  sheet.setConditionalFormatRules(rules.concat(newRules));
};

function hideOldRows(sheet) {
  const { recentDate } = findDateRange();
  if (!recentDate) return;

  const oneWeekAgo = new Date(recentDate);
  oneWeekAgo.setDate(recentDate.getDate() - 7);

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(`A2:A${lastRow}`).getValues();

  values.forEach((date, i) => {
    const rowDate = new Date(date[0]);
    if (rowDate < oneWeekAgo || rowDate > recentDate) {
      sheet.hideRows(i + 2);
    }
  });
};

function hideUnsortedNomoItems(sheet) {
  const lastRow = sheet.getLastRow();
  const valuesA = sheet.getRange(`A2:A${lastRow}`).getValues();
  const valuesB = sheet.getRange(`B2:B${lastRow}`).getValues();
  const valuesC = sheet.getRange(`C2:C${lastRow}`).getValues();
  const valuesD = sheet.getRange(`D2:D${lastRow}`).getValues();
  const quantityValues = sheet.getRange(`D2:D${lastRow}`).getValues();
  const rowsToHide = new Set();

  valuesA.forEach((valueA, i) => {
    const valueAString = valueA[0].toString().toLowerCase();
    if ((valueAString.startsWith("x ")) || (valueAString.startsWith("00"))) {
      rowsToHide.add(i + 2);
    }
  });
  valuesB.forEach((valueB, i) => {
    const valueBString = valueB[0].toString().toLowerCase();
    if ((!rowsToHide.has(i + 2) && valueBString.startsWith("x ")) ||  (valueBString.startsWith("00"))) {
      rowsToHide.add(i + 2);
    }
  });
  valuesC.forEach((valueC, i) => {
    const valueCString = valueC[0].toString().toLowerCase();
    if ((!rowsToHide.has(i + 2) && valueCString.startsWith("x ")) ||  (valueCString.startsWith("00"))) {
      rowsToHide.add(i + 2);
    }
  });
  valuesD.forEach((valueD, i) => {
    const valueDString = valueD[0].toString().toLowerCase();
    if ((!rowsToHide.has(i + 2) && valueDString.startsWith("x ")) ||  (valueDString.startsWith("00"))) {
      rowsToHide.add(i + 2);
    }
  });
  quantityValues.forEach((value, i) => {
    if (value[0] < 1) {
      rowsToHide.add(i + 2);
    }
  });

  rowsToHide.forEach(row => sheet.hideRows(row));
};

function altProcessLiquidSalesData(data, sheet) {
  const filteredData = data.filter(item => item.category === "Mighty Fine");
  const summary = new Map();

  filteredData.forEach(item => {
    const key = `${item.itemName}-${item.modifier}`;
    if (!summary.has(key)) {
      summary.set(key, { itemName: item.itemName, modifier: item.modifier, quantity: 0 });
    }
    summary.get(key).quantity += item.quantity;
  });

  const rows = Array.from(summary.values()).map(item => [item.itemName, item.modifier, item.quantity]);
  const headers = [["Item", "Modifiers", "Qty"]];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  };

  sheet.getDataRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).setHorizontalAlignment('center');
  sheet.setColumnWidth(3, 50);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(1, 200);

  sheet.getDataRange().createFilter();
  sheet.getRange('B1').activate().getFilter().sort(2, true);
  sheet.getRange('A1').activate().getFilter().sort(1, true);

  const formatRange = sheet.getRange("B2:B");
  const rules = sheet.getConditionalFormatRules();
  const formattingRule = SpreadsheetApp.newConditionalFormatRule().whenTextContains("30").setBold(true).setBackground("#e2e2e2").setRanges([formatRange]).build();
  sheet.setConditionalFormatRules(rules.concat(formattingRule));
};

function updateProductLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName("Sales");

  const data = salesSheet.getDataRange().getValues();
  const salesData = data.slice(1).map(row => ({
    date: row[0],                // Column A
    category: String(row[3]),    // Column D
    itemName: String(row[4]),    // Column E
    quantity: row[5],            // Column F
    variation: String(row[6]),   // Column G
    modifier: String(row[8]),    // Column I
    location: String(row[19])    // Column T
  }));

  createDateNamedSheet();
  
  // Venice
  createOrReplaceSheet("Venice List");
  processList(salesData, "Mighty Fine Flavors", "Venice List", "variation", true);
  formatList(ss.getSheetByName("Venice List"));

  // Venice Liquid
  createOrReplaceSheet("Venice Liquid List");
  processLiquidSalesData(salesData, ss.getSheetByName("Venice Liquid List"));
  formatLiquidList(ss.getSheetByName("Venice Liquid List"));

  // North Port
  createOrReplaceSheet("North Port List");
  processList(salesData, "Mighty Fine Vape & Smoke | North Port", "North Port List", "variation", true);
  formatList(ss.getSheetByName("North Port List"));

  // Alternate Venice Liquid
  createOrReplaceSheet("Alt Venice Liquid List");
  altProcessLiquidSalesData(salesData, ss.getSheetByName("Alt Venice Liquid List"));
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Custom Menu')
    .addItem('Make the lists!', 'updateProductLists')
    .addItem('Compare the lists!', 'compareLists')
    .addToUi();
};
