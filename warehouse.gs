function addWarehouse(e, sheet) {
  const name = e.parameter.warehouseName;
  const capacity = e.parameter.warehouseCapacity;
  const address = e.parameter.warehouseAddress;

  if (!name || !capacity || !address) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "All fields required" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Auto-generate WarehouseId (last row id + 1)
  const lastRow = sheet.getLastRow();
  let newId = 1;
  if (lastRow >= 2) {
    const lastId = sheet.getRange(lastRow, 1).getValue(); // Column A = WarehouseId
    newId = Number(lastId) + 1;
  }

  sheet.appendRow([newId, name, capacity, address]);

  return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Warehouse added", id: newId }))
    .setMimeType(ContentService.MimeType.JSON);
}

function editWarehouse(e, sheet) {
  const originalId = e.parameter.originalWarehouseId;
  const newName = e.parameter.newWarehouseName;
  const newCapacity = e.parameter.newWarehouseCapacity;
  const newAddress = e.parameter.newWarehouseAddress;

  if (!originalId) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "WarehouseId required" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(originalId)) {
      sheet.getRange(i + 1, 2).setValue(newName);
      sheet.getRange(i + 1, 3).setValue(newCapacity);
      sheet.getRange(i + 1, 4).setValue(newAddress);

      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Warehouse updated" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Warehouse not found" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function deleteWarehouse(e, sheet) {
  const id = e.parameter.warehouseId;
  if (!id) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "WarehouseId required" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Warehouse deleted" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Warehouse not found" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllWarehouses(sheet) {
  const data = sheet.getDataRange().getValues();
  const warehouses = [];

  for (let i = 1; i < data.length; i++) {
    warehouses.push({
      warehouseId: data[i][0],
      warehouseName: data[i][1],
      warehouseCapacity: data[i][2],
      warehouseAddress: data[i][3],
    });
  }

  return ContentService.createTextOutput(JSON.stringify({ success: true, data: warehouses }))
    .setMimeType(ContentService.MimeType.JSON);
}
