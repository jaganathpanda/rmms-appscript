function submitTransitPass(e,sheet) {
  try {
   // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("transitPass");
    if (!sheet) throw new Error("TransitPass sheet not found");

    const lastRow = sheet.getLastRow();
    let nextId = "transit1";

    if (lastRow > 1) {
      const lastId = sheet.getRange(lastRow, 1).getValue(); // Assuming ID is in Column A
      const match = lastId.match(/transit(\d+)/);
      const nextNum = match ? parseInt(match[1]) + 1 : 1;
      nextId = `transit${nextNum}`;
    }

    const timeStamp = new Date();

    const data = [
      nextId,
      timeStamp,
      e.parameter.ppc,
      e.parameter.miller,
      e.parameter.transitPassNo,
      e.parameter.transitPassDate,
      e.parameter.vehicleNo,
      e.parameter.driverName,
      e.parameter.bag,
      e.parameter.quantity,
      e.parameter.delay,
      e.parameter.acceptedDate
    ];

    sheet.appendRow(data);

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, message: "Transit pass submitted successfully", recordId: nextId })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllTransitPass(sheet1) {
  const sheet = rmmPaddyPurchaseForm.getSheetByName("transitPass");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data = rows.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify({ success: true, data })).setMimeType(ContentService.MimeType.JSON);
}
function updateTransitPass(e,sheet) {
  //const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TransitPass");
  const idToUpdate = e.parameter.transitId;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idToUpdate) {
      const updatedRow = [
        idToUpdate,
        data[i][1], // Keep original timestamp
        e.parameter.ppc,
        e.parameter.miller,
        e.parameter.transitPassNo,
        e.parameter.transitPassDate,
        e.parameter.vehicleNo,
        e.parameter.driverName,
        e.parameter.bag,
        e.parameter.quantity,
        e.parameter.delay,
        e.parameter.acceptedDate
      ];
      sheet.getRange(i + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Updated" })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Record not found" })).setMimeType(ContentService.MimeType.JSON);
}
function deleteTransitPass(e,sheet) {
  //const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TransitPass");
  const idToDelete = e.parameter.transitId;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idToDelete) {
      sheet.deleteRow(i + 1);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Deleted" })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Record not found" })).setMimeType(ContentService.MimeType.JSON);
}