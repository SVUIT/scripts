function RenameFolder() {
  var rootFolderName = "Test"; // root Folder: Test --> Docs
  var rootFolder = DriveApp.getFoldersByName(rootFolderName).next();
  var subFolders = rootFolder.getFolders();

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    var subFolderName = subFolder.getName();
    var subSubFolders = subFolder.getFolders();

    while (subSubFolders.hasNext()) {
      var subSubFolder = subSubFolders.next();
      var subSubFolderName = subSubFolder.getName();

      switch (subSubFolderName) {
        case "K16": // old name
          subSubFolder.setName("2020"); // new name
          break;
        case "K17":
          subSubFolder.setName("2021");
          break;
        case "K18":
          subSubFolder.setName("2022");
          break;
        default:
          continue;
      }
    }
  }
}
