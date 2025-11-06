function commonResponse(passOrFailStatus ,statusMessage ,statusCode  ) {
   var response = {};
  if(passOrFailStatus){
      response = {
        status: "Success",
        message: statusMessage,
        statusCode: statusCode
      };
  }else{
    response = {
        status: "Fail",
        message: statusMessage,
        statusCode: statusCode
      };

  }
  return response;
   
}



function uploadVoucherReceiptImageToDrive(parameters) {
    var dropbox = "voucherReceipt"+parameters.typeOfGoods;
    var fileUrl = "";
    var currentTimeStamp = Date.now();
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    if (folders.hasNext()) {
        folder = folders.next();
    } else {
        folder = DriveApp.createFolder(dropbox);
    }
    var fileName = parameters.vendorName.replace(/ /g, "") + currentTimeStamp +parameters.typeOfGoods+"salse_voucher_receipt.jpg";
    try {
        var contentType = "image/jpg",
            bytes = Utilities.base64Decode(parameters.voucherReceiptImage),
            blob = Utilities.newBlob(bytes, contentType, fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var fileId = file.getId();
        fileUrl = "https://drive.google.com/uc?export=view&id=" + fileId;

        if(parameters.vendorEmail != ""){
            sendGoodsSalesVoucherReceiptEmail(parameters,fileId)
        }
    } catch (error) {
        GmailApp.sendEmail("panda.jaganath@gmail.com", "error on upload image "+parameters.receiptImage, error);
        Logger.log('Error: ' + error);
    }
    return fileUrl;

}

function coulmnSetAsPlainText(sheetName,columnNumber){
  var selectedSheet = rmmPaddyPurchaseForm.getSheetByName(sheetName);
  var lastRow = selectedSheet.getLastRow();
  var range = selectedSheet.getRange(columnNumber+"1:"+columnNumber + lastRow);
  range.setNumberFormat('@');
}




function uploadImageToDrive(parameters) {
    var dropbox = "testFolderImageUpload";
    var fileUrl = "";
    var currentTimeStamp = Date.now();
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    if (folders.hasNext()) {
        folder = folders.next();
    } else {
        folder = DriveApp.createFolder(dropbox);
    }
    var fileName = parameters.farmerName.replace(/ /g, "") + currentTimeStamp + "paddy_receipt.jpg";
    try {
        var contentType = "image/jpg",
            bytes = Utilities.base64Decode(parameters.receiptImage),
            blob = Utilities.newBlob(bytes, contentType, fileName);
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var fileId = file.getId();
        fileUrl = "https://drive.google.com/uc?export=view&id=" + fileId;

        if (parameters.farmerEmail != "") {
            sendPaddyPurchaseReceiptEmail(parameters, fileId)
        }
    } catch (error) {
        GmailApp.sendEmail("panda.jaganath@gmail.com", "error on upload image " + parameters.receiptImage, error);
        Logger.log('Error: ' + error);
    }
    return fileUrl;

}

function sendPaddyPurchaseReceiptEmail(parameters,fileId) {
  var file = DriveApp.getFileById(fileId).getAs(MimeType.JPEG);
  var recipient = parameters.farmerEmail;
  var subject = "Paddy Purchase Receipt - Order-"+parameters.serialNo;
  var perKgPrice = (parameters.perBagPrice/77).toFixed(2);
  var totalPaddayAmount = (parameters.totalPaddyInKg * perKgPrice).toFixed(2) ;

  // Customize your receipt content using HTML
  var body = "<html><body>" +
             "<p>Dear " + parameters.farmerName + ",</p>" +
             "<p>Thank you for your Sale on " + parameters.purchaseDate + ". Below is your sale receipt:</p>" +
             "<table border='1'>" +
             "<tr><th>Serial No.</th><th>Total Paddy(In KG)</th><th>Per Bag Price</th><th>Total</th></tr>" +
             "<tr><td>" + parameters.serialNo + "</td><td>" + parameters.totalPaddyInKg+ "</td><td>" + parameters.perBagPrice + "</td><td>" + totalPaddayAmount + "</td></tr>" +
             "</table>" +
             "<p>Total Amount: " + totalPaddayAmount + "</p>" +
             "<p>If you have any questions, please contact us.</p>" +
             "<p>Thank you,<br>Your Company</p>" +
             "</body></html>";



  var options = {
    htmlBody: body,
    attachments: [file]
  };
   
  // Send HTML-formatted email
  GmailApp.sendEmail(recipient, subject , body , options);
}



function sendGoodsSalesVoucherReceiptEmail(parameters,fileId) {
  var file = DriveApp.getFileById(fileId).getAs(MimeType.JPEG);
  var recipient = parameters.vendorEmail;
  var subject = "Sale Receipt - Order-"+parameters.voucherSerialNo;
  var totalGoodsAmount = (parameters.totalGoodsInKg * parameters.goodsPerKgPrice).toFixed(2) ;
  var totalGoodsBag = (parameters.totalGoodsInKg / parameters.bagPerKg).toFixed(2) ;

  // Customize your receipt content using HTML
  var body = "<html><body>" +
             "<p>Dear " + parameters.vendorName + ",</p>" +
             "<p>Thank you for your Order on " + parameters.saleDate + ". Below is your order receipt:</p>" +
             "<table border='1'>" +
             "<tr><th>Serial No.</th><th>Type Of Goods</th><th>Total Goods(In KG)</th><th>Number Of Bags</th><th>Total Amount</th></tr>" +
             "<tr><td>" + parameters.voucherSerialNo + "</td><td>" + parameters.typeOfGoods + "</td><td>"+ parameters.totalGoodsInKg+ "</td><td>" + totalGoodsBag + "</td><td>" + totalGoodsAmount + "</td></tr>" +
             "</table>" +
             "<p>Total Amount: " + totalGoodsAmount + "</p>" +
             "<p>If you have any questions, please contact us.</p>" +
             "<p>Thank you,<br>Your Company</p>" +
             "</body></html>";



  var options = {
    htmlBody: body,
    attachments: [file]
  };
   
  // Send HTML-formatted email
  GmailApp.sendEmail(recipient, subject , body , options);
}