function paddyPurchaseRecordDelete(e, sheet) {
    var parameters = e.parameter;
    var data = sheet.getDataRange().getValues();
    var paddyRecordId = parameters.paddyRecordId;
    for (var i = 1; i < data.length; i++) {
        var storedPaddyRecordId = data[i][0];
        if (storedPaddyRecordId === paddyRecordId) {
            sheet.getRange(i + 1, 19).setValue("FALSE");
        }

    }
    var response = commonResponse(true, "Paddy Record Deleted Successfully", 200);
    return ContentService
        .createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.TEXT);

}

function paddyPurchaseReport(e, sheet) {
    var mergedArray = new Array(2);
    var sheet = rmmPaddyPurchaseForm.getSheetByName("paddyPurchaseForm");
    var data = sheet.getDataRange().getValues();
    var sheetPayment = rmmPaddyPurchaseForm.getSheetByName("payment");
    var dataPayment = sheetPayment.getDataRange().getValues();
    var rememberFarmerName = Boolean(e.parameter.rememberFarmerName);
    var filteredArray = {};
    if (rememberFarmerName) {
        filteredArray = data.filter(innerArray => innerArray[18] !== false && innerArray[20] == rememberFarmerName);
    } else {
        filteredArray = data.filter(innerArray => innerArray[18] !== false);
    }
    var searchValues = ["PADDY-PAYMENT"];
    var paymentFilteredArray = dataPayment.filter(function(innerArray) {
    return searchValues.some(function(value) {
        return innerArray.includes(value);
    });
  });
  mergedArray[0]=filteredArray;
  mergedArray[1]=paymentFilteredArray;
    var response = commonResponse(true, mergedArray, 200);
    return ContentService
        .createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.TEXT);

}

function addPPC(e, settingSheet) {
  const newPPC = (e.parameter.ppcName || "").trim();
  if (!newPPC) {
    return respond(false, "Missing PPC value.", 404);
  }

  // Get existing PPCs in column A (starting from row 2)
  const lastRow = settingSheet.getLastRow();
  const existingValues = lastRow > 1
    ? settingSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()
    : [];

  // Check for duplicate (case-insensitive)
  const isDuplicate = existingValues.some(v => v.toLowerCase() === newPPC.toLowerCase());
  if (isDuplicate) {
    return respond(false, "PPC already exists.", 403);
  }

  // Append new PPC only to column A
  settingSheet.getRange(lastRow + 1, 1).setValue(newPPC);
  return respond(true, "PPC added successfully.", 200);
}

// Helper function to structure response
function respond(success, message, status) {
  const response = commonResponse(success, message, status);
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllPPCs(sheet) {
  const lastRow = sheet.getLastRow();
   let responseMessage ="";
  // Return empty array if there are no values beyond the header
  if (lastRow < 2) {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .flat()
    .filter(Boolean); // Remove empty or blank entries
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    data: data
  })).setMimeType(ContentService.MimeType.JSON);
}
function deletePPC(e, sheet) {
  const ppcToDelete = e.parameter.ppcName;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toLowerCase() === ppcToDelete.toLowerCase()) {
      sheet.deleteRow(i + 2);
      return jsonResponse(true, "PPC deleted", 200);
    }
  }
  return jsonResponse(false, "PPC not found", 404);
}
function editPPC(e, sheet) {
  const original = e.parameter.originalPPC;
  const updated = e.parameter.newPPC;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toLowerCase() === original.toLowerCase()) {
      sheet.getRange(i + 2, 1).setValue(updated);
      return jsonResponse(true, "PPC updated", 200);
    }
  }
  return jsonResponse(false, "Original PPC not found", 404);
}

function paddyPurchaseForm(e, sheet) {
    var parameters = e.parameter;
    var currentTimeStamp = new Date().toLocaleString();
    var paddyRecordId = "paddyRecordId" + sheet.getLastRow();
    var rmmUserId = parameters.rmmUserId;
    var serialNo = parameters.serialNo;
    var farmerName = parameters.farmerName;
    var rememberFarmerName = parameters.rememberFarmerName;
    var farmerAddress = parameters.farmerAddress;
    var purchaseDate = parameters.purchaseDate;
    var totalPaddyInKg = parameters.totalPaddyInKg;
    var totalPaddyBag = (totalPaddyInKg / 77).toFixed(2);
    var receiptImagePath = "";
    var anyRemarks = parameters.anyRemarks;
    var perKgPrice = (parameters.perBagPrice / 77).toFixed(2);
    var perBagPrice = parameters.perBagPrice;
    var farmeremail = parameters.farmerEmail;
    var farmerMobile = parameters.farmerMobile;
    var typeOfPaddy = parameters.typeOfPaddy;
    var paddyCollectFrom = parameters.paddyCollectFrom;
    var totalPaddayAmount = (totalPaddyInKg * perKgPrice).toFixed(2);
    var receiptImage = parameters.receiptImage;
    var rawDistribution = parameters.warehouseDistribution;
    var response = {};
    if (e.parameter.action === "paddyPurchaseAction") {
        if (receiptImage != "" && receiptImage != null && receiptImage != undefined) {
            var receiptImagePath = uploadImageToDrive(parameters);
        }
        sheet.appendRow([paddyRecordId, currentTimeStamp, rmmUserId, serialNo, farmerName, farmerAddress, farmeremail, farmerMobile, purchaseDate, totalPaddyBag, totalPaddyInKg, perBagPrice, perKgPrice, totalPaddayAmount, receiptImagePath, typeOfPaddy, paddyCollectFrom, anyRemarks, "TRUE", 0, rememberFarmerName]);
        coulmnSetAsPlainText("paddyPurchaseForm", "I");
        updateInventorySheet(paddyRecordId, rawDistribution, "NEW");
        response = commonResponse(true, "Paddy Successfuly Entered ", 200);
    } else if (e.parameter.action === "paddyPurchaseUpdateAction") {
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            var storedPaddyRecordId = data[i][0];
            if (e.parameter.rmmPaddyRecordId === storedPaddyRecordId) {
                sheet.getRange(i + 1, 2).setValue(currentTimeStamp);
                sheet.getRange(i + 1, 3).setValue(rmmUserId);
                sheet.getRange(i + 1, 4).setValue(serialNo);
                sheet.getRange(i + 1, 5).setValue(farmerName);
                sheet.getRange(i + 1, 6).setValue(farmerAddress);
                sheet.getRange(i + 1, 7).setValue(farmeremail);
                sheet.getRange(i + 1, 8).setValue(farmerMobile)
                sheet.getRange(i + 1, 9).setValue(purchaseDate);
                sheet.getRange(i + 1, 10).setValue(totalPaddyBag);
                sheet.getRange(i + 1, 11).setValue(totalPaddyInKg);
                sheet.getRange(i + 1, 12).setValue(perBagPrice);
                sheet.getRange(i + 1, 13).setValue(perKgPrice);
                sheet.getRange(i + 1, 14).setValue(totalPaddayAmount);
                if (receiptImage != "" && receiptImage != null && receiptImage != undefined) {
                    var receiptImagePath = uploadImageToDrive(parameters);
                    sheet.getRange(i + 1, 15).setValue(receiptImagePath);
                }
                sheet.getRange(i + 1, 16).setValue(typeOfPaddy);
                sheet.getRange(i + 1, 17).setValue(paddyCollectFrom);
                sheet.getRange(i + 1, 18).setValue(anyRemarks);
                sheet.getRange(i + 1, 21).setValue(rememberFarmerName);
                updateInventorySheet(storedPaddyRecordId, rawDistribution, "UPDATE");
                response = commonResponse(true, "Record Updated For " + storedPaddyRecordId, 200);
            }
        }
    } else {
        response = commonResponse(false, "Record Not found ", 404);
    }


    return ContentService
        .createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.TEXT);

}