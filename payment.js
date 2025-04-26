function receiveAmountAgainstGoods(e) {
   var parameters = e.parameter;
  var salesVoucherRecordSheet = rmmPaddyPurchaseForm.getSheetByName("salesVoucherRecord");
  var paymentSheet = rmmPaddyPurchaseForm.getSheetByName("payment");
  var currentTimeStamp = new Date().toLocaleString();
  var paymentId = "paymentId"+paymentSheet.getLastRow();
  var data = salesVoucherRecordSheet.getDataRange().getValues();
  var saleVoucherId = parameters.saleVoucherId;
  var rmmUserId = parameters.rmmUserId;
  var padiAmount = parameters.paidAmount;
  var paymentDate = parameters.paymentDate;
  
  var paymentType = parameters.paymentType;
  var anyRemarks = parameters.anyRemarks;
  var filteredArray = data.filter(innerArray => innerArray[0] === saleVoucherId);
  var columnNumber = filteredArray[0][0].split("saleVoucherId")[1];
  if(filteredArray[0][20] == ""){
    salesVoucherRecordSheet.getRange(parseInt(columnNumber) + 1, 21).setValue(padiAmount);
  }else{
      var previousPaidAmount =  parseInt(filteredArray[0][20]);
      
      salesVoucherRecordSheet.getRange(parseInt(columnNumber) + 1, 21).setValue(previousPaidAmount+parseInt(padiAmount))
  }
  paymentSheet.appendRow([paymentId,currentTimeStamp,rmmUserId,saleVoucherId,filteredArray[0][3],"RECIVED",padiAmount,paymentType,paymentDate,anyRemarks]);
  coulmnSetAsPlainText("payment","I");
  response = commonResponse(true,"Payment Successfuly Entered ", 200);
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.TEXT);
}

function payAmountAgainstPaddy(e) {
   var parameters = e.parameter;
  var paddyPurchaseFormSheet = rmmPaddyPurchaseForm.getSheetByName("paddyPurchaseForm");
  var paymentSheet = rmmPaddyPurchaseForm.getSheetByName("payment");
  var currentTimeStamp = new Date().toLocaleString();
  var paymentId = "paymentId"+paymentSheet.getLastRow();
  var data = paddyPurchaseFormSheet.getDataRange().getValues();
  var paddyRecordId = parameters.saleVoucherId;
  var rmmUserId = parameters.rmmUserId;
  var padiAmount = parameters.paidAmount;
  var paymentDate = parameters.paymentDate;
  var paymentType = parameters.paymentType;
  var anyRemarks = parameters.anyRemarks;
  var filteredArray = data.filter(innerArray => innerArray[0] === paddyRecordId);
  var columnNumber = filteredArray[0][0].split("paddyRecordId")[1];
  if(filteredArray[0][19] == ""){
      paddyPurchaseFormSheet.getRange(parseInt(columnNumber) + 1, 20).setValue(padiAmount);
  }else{
    var previousPaidAmount =  parseInt(filteredArray[0][19]);
      
      paddyPurchaseFormSheet.getRange(parseInt(columnNumber) + 1, 20).setValue(previousPaidAmount+parseInt(padiAmount))
  }
  paymentSheet.appendRow([paymentId,currentTimeStamp,rmmUserId,paddyRecordId,"PADDY-PAYMENT","PAID",padiAmount,paymentType,paymentDate,anyRemarks]);
  coulmnSetAsPlainText("payment","I");
  response = commonResponse(true,"Payment Successfuly Entered ", 200);
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.TEXT);
}