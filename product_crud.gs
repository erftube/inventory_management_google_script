function addProduct(product) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products");
  if (!sheet) throw new Error("Products sheet not found");

  const data = sheet.getDataRange().getValues();
  if (data.length < 1) throw new Error("Sheet has no headers");

  const headers = data[0];
  const COL = {};
  headers.forEach((h, i) => {
    if (h) COL[h.toString().trim().toLowerCase()] = i;
  });

  // REQUIRED fields
  const required = ["sku", "product", "category", "cost", "price", "stock", "reorder_threshold"];
  required.forEach(key => {
    if (!(key in product)) {
      throw new Error(`Missing required field: ${key}`);
    }
  });

  // Prevent duplicate SKU
  const skuCol = COL["sku"];
  if (skuCol === undefined) throw new Error("SKU column missing");

  const existingSkus = data.slice(1).map(r => r[skuCol]);
  if (existingSkus.includes(product.sku)) {
    throw new Error(`SKU already exists: ${product.sku}`);
  }

  // Build row in correct column order
  const row = new Array(headers.length).fill("");

  Object.keys(product).forEach(key => {
    const colIndex = COL[key.toLowerCase()];
    if (colIndex !== undefined) {
      row[colIndex] = product[key];
    }
  });

  sheet.appendRow(row);
};



function deleteProductBySku(sku) {
  if (!sku) throw new Error("SKU is required");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products");
  if (!sheet) throw new Error("Products sheet not found");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error("No products to delete");

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  const headers = data[0];
  let skuCol = -1;

  headers.forEach((h, i) => {
    if (h && h.toString().trim().toLowerCase() === "sku") {
      skuCol = i;
    }
  });

  if (skuCol === -1) throw new Error("SKU column not found");

  // Find row index (1-based for Sheets)
  let rowToDelete = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][skuCol] === sku) {
      rowToDelete = i + 1;
      break;
    }
  }

  if (rowToDelete === -1) {
    throw new Error(`SKU not found: ${sku}`);
  }

  sheet.deleteRow(rowToDelete);
};



function updateProductBySku(sku, updates) {
  if (!sku) throw new Error("SKU is required");
  if (!updates || typeof updates !== "object") {
    throw new Error("Updates object is required");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products");
  if (!sheet) throw new Error("Products sheet not found");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error("No products to update");

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Map headers
  const headers = data[0];
  const COL = {};
  headers.forEach((h, i) => {
    if (h) COL[h.toString().trim().toLowerCase()] = i;
  });

  if (COL["sku"] === undefined) throw new Error("SKU column not found");

  // Find product row
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL["sku"]] === sku) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error(`SKU not found: ${sku}`);
  }

  // Apply updates
  const row = data[rowIndex].slice(); // clone row

  Object.keys(updates).forEach(key => {
    const colIndex = COL[key.toLowerCase()];
    if (colIndex === undefined) {
      throw new Error(`Unknown column: ${key}`);
    }

    // Prevent SKU change unless you explicitly allow it
    if (key.toLowerCase() === "sku") {
      throw new Error("SKU cannot be changed");
    }

    row[colIndex] = updates[key];
  });

  // Write back updated row
  sheet.getRange(rowIndex + 1, 1, 1, lastCol).setValues([row]);
};



