
function getDashboardStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Products");
  const data = sheet.getDataRange().getValues();
  
  // Map column names to indexes
  const headers = data[0];
  const COL = {};
  headers.forEach((name, i) => {
    COL[name.toLowerCase()] = i;
  });

  const rows = data.slice(1); // skip header
  let total = 0;
  let low = 0;
  let out = 0;
  let cost = 0;
  let price = 0;

  // Array to hold structured product items
  const items = [];

  rows.forEach(row => {
    const sku = row[COL["sku"]];
    const product = row[COL["product"]];
    const category = row[COL["category"]];
    const costVal = Number(row[COL["cost"]]) || 0;
    const priceVal = Number(row[COL["price"]]) || 0;
    const stock = Number(row[COL["stock"]]) || 0;
    const reorder = Number(row[COL["reorder_threshold"]]) || 0;

    if (!sku) return; // skip empty rows

    total++;

    let status = "In Stock";
    if (stock === 0) {
      status = "Out of Stock";
      out++;
    } else if (stock <= reorder) {
      status = "Low Stock";
      low++;
    }

    cost += stock * costVal;
    price += stock * priceVal;

    items.push({
      sku,
      product,
      category,
      cost: costVal,
      price: priceVal,
      stock,
      reorder,
      status
    });
  });

  return {
    total,
    low,
    out,
    cost,
    price,
    items
  };
}
