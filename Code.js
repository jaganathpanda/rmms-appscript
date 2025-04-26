var rmmPaddyPurchaseForm = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/14MbYxR1XZ9hmNsay6tN23pN_RCIGdWF-FKSfOEEPFbY/edit#gid=1725595892");


function doPost(e) { 
  if( e.parameter.action === "paddyPurchaseAction" || e.parameter.action === "paddyPurchaseUpdateAction"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("paddyPurchaseForm");
      return paddyPurchaseForm(e,sheet);
  }
  if( e.parameter.action === "paddyPurchaseReport"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("paddyPurchaseForm");
      return paddyPurchaseReport(e,sheet);
  }
  if( e.parameter.action === "paddyPurchaseRecordDelete"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("paddyPurchaseForm");
      return paddyPurchaseRecordDelete(e,sheet);
  }
  if( e.parameter.action === "salesVoucherInsertAction" || e.parameter.action === "salesVoucherUpdateAction"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("salesVoucherRecord");
      return salesVoucherInsertUpdateRecord(e,sheet);
  }

    if( e.parameter.action === "salesVoucherDeleteRecord"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("salesVoucherRecord");
      return salesVoucherDeleteRecord(e,sheet);
  }

  if( e.parameter.action === "allSalesVoucherReport"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("salesVoucherRecord");
      return allSalesVoucherReport(e,sheet);
  }

  if( e.parameter.action === "receiveAmountAgainstGoods"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("salesVoucherRecord");
      return receiveAmountAgainstGoods(e);
  }

  if( e.parameter.action === "payAmountAgainstPaddy"){
      var sheet = rmmPaddyPurchaseForm.getSheetByName("salesVoucherRecord");
      return payAmountAgainstPaddy(e);
  }

  if( e.parameter.action === "addEmployee" || e.parameter.action === "editEmployee"){
      return employeeManagement(e);
  }

  if( e.parameter.action === "employeeReport"){
      return employeeReport(e);
  }

  if( e.parameter.action === "addPPC"){
      const settingSheet = rmmPaddyPurchaseForm.getSheetByName("setting");
      return addPPC(e,settingSheet);
  }

  if( e.parameter.action === "editPPC"){
      const settingSheet = rmmPaddyPurchaseForm.getSheetByName("setting");
      return editPPC(e,settingSheet);
  }

  if( e.parameter.action === "deletePPC"){
      const settingSheet = rmmPaddyPurchaseForm.getSheetByName("setting");
      return deletePPC(e,settingSheet);
  }

  if( e.parameter.action === "getAllPPCs"){
     const sheet = rmmPaddyPurchaseForm.getSheetByName("setting");
      return getAllPPCs(sheet);
  }
  if (e.parameter.action === "submitTransitPass") {
    const sheet = rmmPaddyPurchaseForm.getSheetByName("transitPass");
    return submitTransitPass(e,sheet);
  }

  if (e.parameter.action === "getAllTransitPass") {
    const sheet = rmmPaddyPurchaseForm.getSheetByName("transitPass");
    return getAllTransitPass(sheet);
  }

  if (e.parameter.action === "updateTransitPass") {
    const sheet = rmmPaddyPurchaseForm.getSheetByName("transitPass");
    return updateTransitPass(e,sheet);
  }
  if (e.parameter.action === "deleteTransitPass") {
    const sheet = rmmPaddyPurchaseForm.getSheetByName("transitPass");
    return deleteTransitPass(e,sheet);
  }

  if (e.parameter.action === "createRazorpayOrder") {
    //var params = JSON.parse(e.postData.contents);
    var paidAmount = e.parameter.paidAmount; // Amount to be paid
    var paymentDate = e.parameter.paymentDate; // Payment date from frontend
      var result = createRazorpayOrder(paidAmount, paymentDate);
      // If order creation is successful, return order details
      if (result.status === "Success") {
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ status: "Error", message: "Failed to create Razorpay order" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    

  }
  return ContentService.createTextOutput(
    JSON.stringify({ success: false, message: "Invalid action" })
  ).setMimeType(ContentService.MimeType.JSON);

}


