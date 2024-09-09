function RenameFolder() {
  var spreadsheetId = "16u9Xg1c-81VL6uxPTSKJczKfwccSW5UVo9Z9d7wtJho"; // ID của Sheet DSMH.
  var sheetName = "Sheet1"; // Tên page sử dụng trong DSMH.

  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    var parentFolderName = String(data[i][0]).trim(); 
    var oldFolderName = String(data[i][1]).trim(); 
    var newFolderName = String(data[i][2]).trim(); 

    var parentFolder = DriveApp.getFoldersByName(parentFolderName).next();
    var subFolders = parentFolder.getFoldersByName(oldFolderName);
    
    while (subFolders.hasNext()) {
      var subFolder = subFolders.next();
      subFolder.setName(newFolderName);
      Logger.log("Đã đổi tên thư mục " + oldFolderName + " trong thư mục " + parentFolderName + " thành: " + newFolderName); //Log xem quá trình hoạt động.
    }
  }
}
