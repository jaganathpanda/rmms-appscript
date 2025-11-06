function updateInventorySheet(paddyRecordId, rawDistribution, mode) {
  try {
    const ss = rmmPaddyPurchaseForm;
    const sheet = ss.getSheetByName("Inventory");
    if (!sheet) {
      Logger.log("❌ Inventory sheet not found");
      return;
    }

    if (!rawDistribution) {
      Logger.log("⚠️ No warehouseDistribution provided");
      return;
    }

    let distributionData;
    try {
      distributionData = JSON.parse(rawDistribution);
    } catch (err) {
      Logger.log("❌ Invalid JSON in warehouseDistribution: " + err);
      return;
    }

    if (!Array.isArray(distributionData) || distributionData.length === 0) {
      Logger.log("⚠️ Empty warehouseDistribution array");
      return;
    }

    if (mode === "UPDATE") {
      // Remove old entries for this paddy purchase
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i > 0; i--) {
        if (data[i][1] === paddyRecordId) {
          sheet.deleteRow(i + 1);
        }
      }
    }

    const nextIdStart = sheet.getLastRow();
    let nextNum = nextIdStart > 1 ? nextIdStart : 1;

    distributionData.forEach((entry) => {
      const inventoryId = "INV" + (nextNum + 1);
      const warehouseId = entry.warehouseId || "";
      const paddyTypeId = entry.paddyTypeId || "";
      const numberOfBags = Number(entry.bags) || 0;
      sheet.appendRow([inventoryId, paddyRecordId, warehouseId, paddyTypeId, numberOfBags]);
      nextNum++;
    });

    Logger.log(`✅ Inventory ${mode} successful for ${paddyRecordId}`);
  } catch (error) {
    Logger.log("❌ Error in updateInventorySheet: " + error);
  }
}

// ================= INVENTORY FUNCTIONS =====================

// 🔸 Get all inventory records
function getAllInventory(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const records = data.map(row => ({
    inventoryId: row[0],
    paddyPurchaseId: row[1],
    warehouseId: Number(row[2]),
    paddyTypeId: row[3],
    bags: Number(row[4])
  }));

  return ContentService.createTextOutput(JSON.stringify(records))
    .setMimeType(ContentService.MimeType.JSON);
}

// 🔸 Calculate total stock by Paddy Type
function getTotalStock(sheet) {
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove header

  const stock = {};
  data.forEach(row => {
    const paddyType = row[3];
    const bags = Number(row[4]);
    stock[paddyType] = (stock[paddyType] || 0) + bags;
  });

  return ContentService.createTextOutput(JSON.stringify(stock))
    .setMimeType(ContentService.MimeType.JSON);
}

// 🔸 Get full details for a given Paddy Purchase ID (distribution across warehouses)
function getPaddyInventoryDetails(e, inventorySheet, warehouseSheet, paddySheet) {
  const paddyId = e.parameter.paddyId;
  if (!paddyId) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Missing paddyId" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const inventory = getSheetDataAsObjects(inventorySheet);
  const warehouses = getSheetDataAsObjects(warehouseSheet);
  const paddies = getSheetDataAsObjects(paddySheet);

  const paddy = paddies.find(p => p[Object.keys(p)[0]] === paddyId); // find by first column (ID)
  const relatedInventory = inventory.filter(i => i["PADDY PURCHASE ID"] === paddyId);

  const stockDistribution = relatedInventory.map(i => {
    const warehouse = warehouses.find(w => Number(w["Warehouse Id"]) === Number(i["WAREHOUSE ID"]));
    return {
      warehouse: warehouse ? warehouse["WarehouseName"] : "Unknown",
      paddyType: i["PADDYTYPE ID"],
      bags: i["NUMBER OF BAGS"]
    };
  });

  const result = {
    paddy,
    stockDistribution
  };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 🔸 Get all stock stored in a given warehouse
function getWarehouseStock(e, inventorySheet, warehouseSheet) {
  const warehouseId = Number(e.parameter.warehouseId);
  if (!warehouseId) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Missing warehouseId" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const inventory = getSheetDataAsObjects(inventorySheet);
  const warehouses = getSheetDataAsObjects(warehouseSheet);

  const warehouse = warehouses.find(w => Number(w["Warehouse Id"]) === warehouseId);
  const stock = inventory.filter(i => Number(i["WAREHOUSE ID"]) === warehouseId);

  const result = {
    warehouse: warehouse || { error: "Warehouse not found" },
    stock
  };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 🔸 Utility to convert a sheet to an array of objects (header-based)
function getSheetDataAsObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => (obj[header] = row[i]));
    return obj;
  });
}
