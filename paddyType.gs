/**
 * Add a new Paddy Type
 */
function addPaddyType(e, sheet) {
  try {
    const name = e.parameter.paddyTypeName;
    const description = e.parameter.description || "";
    if (!name) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Missing paddyTypeName" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const lastRow = sheet.getLastRow();
    const newId = lastRow >= 2 ? lastRow : 1;
    const nextId = "PT" + (newId + 1);

    sheet.appendRow([nextId, name, description]);

    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Paddy type added successfully" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Fetch all Paddy Types
 */
function getAllPaddyTypes(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const result = data.map(row => ({
      paddyTypeId: row[0],
      paddyTypeName: row[1],
      paddyTypeDescription: row[2]
    }));

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Edit an existing Paddy Type
 */
function editPaddyType(e, sheet) {
  try {
    const id = e.parameter.paddyTypeId;
    const newName = e.parameter.newPaddyTypeName;
    const newDesc = e.parameter.newDescription;

    if (!id) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Missing paddyTypeId" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.getRange(i + 1, 2).setValue(newName || data[i][1]);
        sheet.getRange(i + 1, 3).setValue(newDesc || data[i][2]);

        return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Paddy type updated successfully" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Paddy type not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Delete a Paddy Type
 */
function deletePaddyType(e, sheet) {
  try {
    const id = e.parameter.paddyTypeId;
    if (!id) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Missing paddyTypeId" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Paddy type deleted successfully" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Paddy type not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
