function salesVoucherInsertUpdateRecord(e,sheet){
  var parameters = e.parameter;
  var currentTimeStamp = new Date().toLocaleString();
  var saleVoucherId = "saleVoucherId"+sheet.getLastRow();
  var rmmUserId = parameters.rmmUserId;
  var typeOfGoods = parameters.typeOfGoods;
  var voucherSerialNo = parameters.voucherSerialNo;
  var vendorName = parameters.vendorName;
  var vendorAddress = parameters.vendorAddress;
  var driverName = parameters.driverName;
  var vehicleNumber = parameters.vehicleNumber;
  var vendorEmail = parameters.vendorEmail;
  var vendorPhone = parameters.vendorPhone;
  var saleDate = parameters.saleDate;
  var bagPerKg = parameters.bagPerKg;
  var goodsPerKgPrice = parameters.goodsPerKgPrice;
  var totalGoodsInKg = parameters.totalGoodsInKg;
  var totalGoodsBag = (totalGoodsInKg / bagPerKg).toFixed(2) ;
      e.parameters.totalGoodsBag = totalGoodsBag;
  var totalGoodsAmount = (totalGoodsInKg * goodsPerKgPrice).toFixed(2) ;
  var goodsCategory = parameters.goodsCategory;
  var voucherReceiptImagePath  = "";
  var voucherReceiptImage = parameters.voucherReceiptImage;
  
 
  var response = {};
  if(e.parameter.action === "salesVoucherInsertAction"){
      if(voucherReceiptImage != "" && voucherReceiptImage != null && voucherReceiptImage != undefined){
       var voucherReceiptImagePath = uploadVoucherReceiptImageToDrive(parameters);
     }
  sheet.appendRow([saleVoucherId,currentTimeStamp,rmmUserId,typeOfGoods,voucherSerialNo,vendorName,vendorAddress,driverName,vehicleNumber,vendorEmail,vendorPhone,saleDate,bagPerKg,totalGoodsInKg,goodsPerKgPrice,totalGoodsBag,totalGoodsAmount,goodsCategory,voucherReceiptImagePath,"TRUE",0]);
  coulmnSetAsPlainText("salesVoucherRecord","L")
  response = commonResponse(true,"Voucher Successfuly Entered ", 200);

  }else if(e.parameter.action === "salesVoucherUpdateAction"){
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var storedRmmSalesVoucherId = data[i][0];
         if (e.parameter.rmmSalesVoucherId === storedRmmSalesVoucherId) {
           sheet.getRange(i + 1, 2).setValue(currentTimeStamp);
           sheet.getRange(i + 1, 3).setValue(rmmUserId);
           sheet.getRange(i + 1, 4).setValue(typeOfGoods);
           sheet.getRange(i + 1, 5).setValue(voucherSerialNo);
           sheet.getRange(i + 1, 6).setValue(vendorName);
           sheet.getRange(i + 1, 7).setValue(vendorAddress);
           sheet.getRange(i + 1, 8).setValue(driverName)
           sheet.getRange(i + 1, 9).setValue(vehicleNumber);
           sheet.getRange(i + 1, 10).setValue(vendorEmail);
           sheet.getRange(i + 1, 11).setValue(vendorPhone);
           sheet.getRange(i + 1, 12).setValue(saleDate);
           sheet.getRange(i + 1, 13).setValue(bagPerKg);
           sheet.getRange(i + 1, 14).setValue(totalGoodsInKg);
           sheet.getRange(i + 1, 15).setValue(goodsPerKgPrice);
           sheet.getRange(i + 1, 16).setValue(totalGoodsBag);
           sheet.getRange(i + 1, 17).setValue(totalGoodsAmount);
           sheet.getRange(i + 1, 18).setValue(goodsCategory);
           if(voucherReceiptImage != "" && voucherReceiptImage != null && voucherReceiptImage != undefined){
             var voucherReceiptImagePath=uploadVoucherReceiptImageToDrive(parameters);
             sheet.getRange(i + 1, 19).setValue(voucherReceiptImagePath);
           } 
           response = commonResponse(true,"Record Updated For "+storedRmmSalesVoucherId, 200);   
         }
      }
  }else{
      response = commonResponse(false,"Record Not found ", 404);   
  }


return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.TEXT);

}

function salesVoucherDeleteRecord(e,sheet){
  var parameters = e.parameter;
  var flag = false;
  var data = sheet.getDataRange().getValues();
  var salesVocherRecordId = parameters.salesVocherRecordId;
  for (var i = 1; i < data.length; i++) {
    var storedSalesVocherRecordId= data[i][0];
    if(storedSalesVocherRecordId === salesVocherRecordId){
       sheet.getRange(i + 1, 20).setValue("FALSE");
      flag = true;
    }

  }
    var response = flag ? commonResponse(true,"Sales Voucher Record Deleted Successfully", 200) : commonResponse(true,"Sales Voucher Record not found", 404)
  return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.TEXT);

}


function allSalesVoucherReport(e,sheet){
  var mergedArray = new Array(2);
  var sheet = rmmPaddyPurchaseForm.getSheetByName("salesVoucherRecord");
  var sheetPayment = rmmPaddyPurchaseForm.getSheetByName("payment");
  var data = sheet.getDataRange().getValues();
  var dataPayment = sheetPayment.getDataRange().getValues();
  var filteredArray = data.filter(innerArray => innerArray[19] !== false);
  var searchValues = ["RICE", "HUSK","RICE-BRAN","PADDY","BROKEN-RICE"];
  var payentFilteredArray = dataPayment.filter(function(innerArray) {
    return searchValues.some(function(value) {
        return innerArray.includes(value);
    });
  });
  mergedArray[0]=filteredArray;
  mergedArray[1]=payentFilteredArray;
  var response = commonResponse(true,mergedArray, 200);

  return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.TEXT);

}