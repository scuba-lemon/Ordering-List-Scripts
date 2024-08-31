function processOrderingList() {
  let orderingListSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/19rV2sw-JVJaYHYmG0xC5FrpqE9uGjKYKOdBOiwvBxI4/edit?gid=389276120#gid=389276120");
  let sheet = orderingListSpreadsheet.getSheetByName("Ordering List")
  let range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9);
  let data = range.getValues();

  let vList = [];
  let npList = [];

  data.forEach(row => {
    row = row.map((value, index) => (index < 5 ? String(value).toLowerCase() : value));

    if (row[7]) {
      let date = new Date(row[7]);
      row[7] = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd/yy');
    };

    if (row[0] === 'v') {
      vList.push({
        itemName: row[4],
        brand: row[2],
        subCat: row[3],
        category: row[1],
        stock: row[5],
        sold: row[6],
        listDate: row[7],
        orderDate: row[8]
      });
    } else if (row[0] === 'np') {
      npList.push({
        itemName: row[4],
        brand: row[2],
        subCat: row[3],
        category: row[1],
        stock: row[5],
        sold: row[6],
        listDate: row[7],
        orderDate: row[8]
      });
    };
  });

  return { vList, npList };
};

function processLocationSheet(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4);
  let data = range.getValues();

  let processedData = data.map(row => {
    return row.map((value, index) => (index < 4 ? String(value).toLowerCase() : value));
  });

  return processedData;
};

function createOrReplaceSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  ss.insertSheet(sheetName);
};

function writeMatchingResultsToSheet(sheet, data) {
  sheet.getRange(1, 1, 1, 7).setValues([['Category', 'Brand', 'SubCat', 'Item Name', 'Stock', 'Sold L.M.', 'Listed Date']]);
  
  if (data.length > 0) {
    let values = data.map(item => [item.category, item.brand, item.subCat, item.itemName, item.stock, item.sold, item.listDate]);
    sheet.getRange(2, 1, values.length, 7).setValues(values);
  };

  let range = sheet.getDataRange();
  range.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment('center').createFilter();
  sheet.autoResizeColumn(1);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 200);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);
  sheet.autoResizeColumn(7);
};

function writeNonMatchingResultsToSheet(sheet, data) {
  sheet.getRange(1, 1, 1, 5).setValues([['Category', 'Brand', 'SubCat', 'Item', 'Qty']]);

  if (data.length > 0) {
    let values = data.map(item => {
      let brand, subCat;

      if (item.category !=="disposables" && item.itemName.includes('-')) {
        [brand, subCat] = item.itemName.split('-').map(str => str.trim());
        itemName = subCat;
      } else {
        brand = item.itemName;
        subCat = 'n/a';
        itemName = brand;
      };

      if ((item.category === "sstash" && subCat !== "cones" && subCat !== "fluid") || item.category === "accessories") {
        [brand, subCat] = [subCat, brand];
      } else if (item.category === "sstash" && subCat === "cones") {
        item.variationName = brand + " - " + item.variationName;
        brand = "n/a"
      } else if (item.category === "sstash" && subCat === "fluid") {
        item.variationName = brand + " - " + item.variationName;
        subCat = brand;
        brand = "n/a"
      };

      if (item.category === "cbd") {
        if (item.variationName.includes('-')) {
          subCat = item.variationName.split('-')[0].trim();
          item.variationName = item.variationName.split('-')[1].trim();
        };
      };

      if (item.category === "coils & pods") {
        let subCatValue = subCat;
        subCat = brand.split(' ')[1].trim() + " - " + subCatValue;
        brand = brand.split(' ')[0].trim();
      };

      if (item.category === "e-liquid") {
        const variation = item.variationName.toLowerCase();

        if (variation.includes("0mg") || variation.includes("3mg") || variation.includes("6mg")) {
          subCat = "freebase";
        } else {
          subCat = "salt";
        };
      };

      if ((item.category === "kits" || item.category === "mods" || item.category === "mech" || item.category === "tanks") && !brand.startsWith("x ")) {
        brand = brand.split(' ')[0].trim();
      };

      return [item.category, brand, subCat, item.variationName, item.qty];
    });

    sheet.getRange(2, 1, values.length, 5).setValues(values);
  };

  sheet.getDataRange().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setHorizontalAlignment('center')
  sheet.getRange(1, 4, sheet.getLastRow(), 2).setHorizontalAlignment('center');
  sheet.setColumnWidth(5, 50);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(1, 100);
  sheet.getDataRange().createFilter();
  sheet.getRange('D1').activate().getFilter().sort(2, true);
  sheet.getRange('C1').activate().getFilter().sort(3, true);
  sheet.getRange('B1').activate().getFilter().sort(2, true);
  sheet.getRange('A1').activate().getFilter().sort(1, true);
  hideUnsortedNomoItems(sheet);
};

function hideNPVending() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("North Port Non-Matches");
  const lastRow = sheet.getLastRow();
  const valuesA = sheet.getRange(`A2:A${lastRow}`).getValues();
  const rowsToHide = new Set();

  valuesA.forEach((valueA, i) => {
    const valueAString = valueA[0].toString().toLowerCase();
    if (valueAString.startsWith("vending")) {
      rowsToHide.add(i + 2);
    };
  });

  rowsToHide.forEach(row => sheet.hideRows(row));
};

function compareLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let { vList, npList } = processOrderingList();
  let veniceData = processLocationSheet("Venice List");
  let northPortData = processLocationSheet("North Port List");

  let veniceItemNames = new Set(vList.map(item => item.itemName));
  let northPortItemNames = new Set(npList.map(item => item.itemName));

  let veniceMatches = vList.filter(item => 
    veniceData.some(row => 
      row[1] === item.itemName && (
        row[0].includes(item.brand) || row[0].includes(item.subCat)
      )
    )
  );

  let northPortMatches = npList.filter(item => 
    northPortData.some(row => 
      row[1] === item.itemName && (
        row[0].includes(item.brand) || row[0].includes(item.subCat)
      )
    )
  );

  let veniceNotInOrdering = veniceData.filter(row => 
    !veniceItemNames.has(row[1])
  ).map(row => ({
    category: row[2],
    itemName: row[0],
    variationName: row[1],
    qty: row[3]
  }));

  let northPortNotInOrdering = northPortData.filter(row => 
    !northPortItemNames.has(row[1])
  ).map(row => ({
    category: row[2],
    itemName: row[0],
    variationName: row[1],
    qty: row[3]
  }));

  createOrReplaceSheet(" | ");

  createOrReplaceSheet("Venice Matches");
  let veniceMatchSheet = ss.getSheetByName("Venice Matches");
  writeMatchingResultsToSheet(veniceMatchSheet, veniceMatches);

  createOrReplaceSheet("Venice Non-Matches");
  let veniceNonMatchingSheet = ss.getSheetByName("Venice Non-Matches");
  writeNonMatchingResultsToSheet(veniceNonMatchingSheet, veniceNotInOrdering);

  createOrReplaceSheet("North Port Matches");
  let northPortMatchSheet = ss.getSheetByName("North Port Matches");
  writeMatchingResultsToSheet(northPortMatchSheet, northPortMatches);

  createOrReplaceSheet("North Port Non-Matches");
  let npNonMatchingsSheet = ss.getSheetByName("North Port Non-Matches");
  writeNonMatchingResultsToSheet(npNonMatchingsSheet, northPortNotInOrdering);
  hideNPVending();
};
