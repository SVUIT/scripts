function ListingFolder() {
  try {
    var rootFolderId = '1TjIygC_EermjfPRFDxsyw3qa3sOqMA3L'; //ID folder Docs
    
    // zo folder
    var rootFolder = DriveApp.getFolderById(rootFolderId);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear();
    
    sheet.getRange(1, 1).setValue('Parent Folder');
    sheet.getRange(1, 2).setValue('Sub Folder');
    sheet.getRange(1, 3).setValue('Folder Link');
    
    var row = 2;
    var parentFolders = rootFolder.getFolders();
    while (parentFolders.hasNext()) {
      var parentFolder = parentFolders.next();
      var parentFolderName = parentFolder.getName();
      var parentFolderUrl = parentFolder.getUrl();
      
      var subFolders = parentFolder.getFolders();
      
      if (!subFolders.hasNext()) {
        sheet.getRange(row, 1).setValue(parentFolderName);
        sheet.getRange(row, 2).setValue(' ');
        sheet.getRange(row, 3).setValue(parentFolderUrl);
        row++;
      } else {
        var firstSubfolder = true;
        while (subFolders.hasNext()) {
          var subFolder = subFolders.next();
          var subFolderUrl = subFolder.getUrl();
          if (firstSubfolder) {
            sheet.getRange(row, 1).setValue(parentFolderName);
            sheet.getRange(row, 3).setValue(parentFolderUrl);
            firstSubfolder = false;
          }
          sheet.getRange(row, 2).setValue(subFolder.getName());
          sheet.getRange(row, 3).setValue(subFolderUrl);
          row++;
        }
      }
    }
  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}
