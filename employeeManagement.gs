function employeeManagement(e){
  var parameters = e.parameter;
  var sheet = rmmPaddyPurchaseForm.getSheetByName("employeeManagement");
  var currentTimeStamp = new Date().toLocaleString();
  var employeeId = "employeeId"+sheet.getLastRow();
  var rmmUserId = parameters.rmmUserId;
  var employeeName = parameters.employeeName;
  var employeePhoneNumber = parameters.employeePhoneNumber;
  var employeeAge = parameters.employeeAge;
  var employeeAddress = parameters.employeeAddress;
  var employeeRole = parameters.employeeRole ;
  var employeeJoinDate = parameters.employeeJoinDate;
  var employeeSalary = parameters.employeeSalary;
  var employeeStatus = parameters.employeeStatus;
  if(e.parameter.action === "addEmployee"){
      sheet.appendRow([employeeId,rmmUserId,currentTimeStamp,employeeName,employeePhoneNumber,employeeAge,employeeAddress,
      employeeRole,employeeJoinDate,employeeSalary,employeeStatus]);
      coulmnSetAsPlainText("employeeManagement","I");
      response = commonResponse(true,"Employee Successfuly Entered ", 200);
  }else if(e.parameter.action === "editEmployee"){
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var storedEmployeeId = data[i][0];
         if (e.parameter.employeeId === storedEmployeeId) {
           sheet.getRange(i + 1, 2).setValue(rmmUserId);
           sheet.getRange(i + 1, 3).setValue(currentTimeStamp);
           sheet.getRange(i + 1, 4).setValue(employeeName);
           sheet.getRange(i + 1, 5).setValue(employeePhoneNumber);
           sheet.getRange(i + 1, 6).setValue(employeeAge);
           sheet.getRange(i + 1, 7).setValue(employeeAddress);
           sheet.getRange(i + 1, 8).setValue(employeeRole)
           sheet.getRange(i + 1, 9).setValue(employeeJoinDate);
           sheet.getRange(i + 1, 10).setValue(employeeSalary);
           sheet.getRange(i + 1, 11).setValue(employeeStatus);   
           response = commonResponse(true,"Employee Updated For "+employeeId, 200);   
         }
      }
  }else{
      response = commonResponse(false,"Record Not found ", 404);   
  }


return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.TEXT);

}

function employeeReport(e){
  var sheet = rmmPaddyPurchaseForm.getSheetByName("employeeManagement");
  var data = sheet.getDataRange().getValues();
  var response = commonResponse(true,data, 200);
  return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.TEXT);

}