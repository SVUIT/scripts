function InsertToSheet() {
  var folderName = "Docs"; 
  var folder = DriveApp.getFoldersByName(folderName).next();
  
  var sheetName = "DSMH";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(["Tên Thư Mục"]);
  } else {
    sheet.clear();
    sheet.appendRow(["Tên Thư Mục"]);
  }
  
  function processFolder(folder) {
    var subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      var folderName = subfolder.getName();
      sheet.appendRow([folderName]);
      processFolder(subfolder); // Gọi lại hàm cho các thư mục con
    }
  }
  
  processFolder(folder);
}
