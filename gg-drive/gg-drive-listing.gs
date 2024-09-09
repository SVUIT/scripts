/**
* Functions to get Google Drive file sizes and other details
* regarding files and folders.
*
* Files created with Google Editor Apps such as Google Sheets,
* Google Docs and Google Slides always have a size of 0.
* Other types of files will show the number of bytes the
* files take in Google Drive.
*
* @OnlyCurrentDoc
*/

/**
* Simple trigger that runs each time the user opens the
* spreadsheet.
*
* Adds menu items to insert details of Google Drive files
* in spreadsheet cells.
*
* @param {Object} e The onOpen() event object.
*/
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Cù lao time')
    .addItem('Tới lúc cù lao rồi', 'listFileInfoBySubFolder')
    .addSeparator()
    .addItem('List file sizes by file IDs in selected cells', 'listFileSizesByIdsInSelectedCells')
    .addToUi();
}

function clearSheet_() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  activeSheet.getDataRange().clearContent();
}

/**
* Gets the names, types, URLs and sizes of files in a folder and its
* subfolders, and inserts the list in the spreadsheet starting
* at the active cell.
*
* The function will overwrite existing values in the spreadsheet
* in the range that starts at the active cell and extends a
* total of four cells to the right and as many rows down as
* there are results.
*/
function listFileInfoBySubFolder() {
  clearSheet_();
  // version 1.1, written by --Hyde, 15 July 2020
  //  - add recursion
  //  - add file size and file type info
  //  - see https://support.google.com/docs/thread/58260211
  // version 1.0, written by --Hyde, 16 June 2020
  //  - initial version
  //  - see https://support.google.com/docs/thread/53541982
  const activeRange = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1");
  const chosen = chooseFileOrFolder_();
  if (!chosen) {
    return;
  }
  if (chosen.error) {
    showMessage_(chosen.error, 30);
    return;
  }
  let fileList = [['Folder', 'File', 'Type', 'URL', 'Size']];
  switch (chosen.fileType) {
    case 'application/vnd.google-apps.folder':
      try {
        let folder = DriveApp.getFolderById(chosen.fileId);
        fileList = fileList.concat(getFileInfoBySubFolder_(folder));
      } catch (error) {
        showAndThrow_(error);
      }
      break;
    default:
      fileList.push([chosen.file.getParents().next().getName(), chosen.file.getName(), chosen.fileType, chosen.file.getUrl(), chosen.file.getSize()]);
  }
  activeRange.offset(0, 0, fileList.length, fileList[0].length).setValues(fileList);
}

/**
* Retrieves a list of file and folder IDs in the currently
* selected cells in the spreadsheet, and puts the sizes of the
* files in the cells next to the selected cells.
*
* The function will overwrite existing values in the spreadsheet
* in the cells next to the selected cells without warning.
*/
function listFileSizesByIdsInSelectedCells() {
  // version 1.0, written by --Hyde, 15 July 2020
  //  - initial version
  //  - see https://support.google.com/docs/thread/53541982
  var fileIdRange = SpreadsheetApp.getActiveRange();
  var fileSizeRange = fileIdRange.offset(0, 1);
  var fileIds = fileIdRange.getDisplayValues();
  if (fileIds[0].length !== 1 || (fileIds.length === 1 && !fileIds[0][0])) {
    showMessage_('Please select exactly one column that contains file or folder IDs.', 10);
    return;
  }
  var sizes = fileIds.map(row => row[0]).map(function (fileId) {
    if (!String(fileId).trim()) {
      return [null];
    }
    let file;
    try {
      file = DriveApp.getFileById(fileId);
    } catch (error) {
      return ['Invalid file or folder ID "' + fileId + '".'];
    }
    let fileType = file.getMimeType();
    switch (fileType) {
      case 'application/vnd.google-apps.folder':
        try { // get the sum of file sizes in folder
          let folder = DriveApp.getFolderById(fileId);
          let fileList = getFileInfoBySubFolder_(folder);
          let lastColumn = fileList[0].length - 1;
          return [fileList[0][lastColumn]];
        } catch (error) {
          return [error.message];
        }
        break;
      default:
        return [file.getSize()];
    }
  });
  fileSizeRange.setValues(sizes);
}

/**
* Gets the names, types, URLs and sizes of files in a folder and its
* subfolders recursively.
*
* The size reported for each folder only represents the sum of
* sizes of the files stored directly in the folder, and does not
* include the size of the files stored in its subfolders. The 
* size of each subfolder is reported separately in a similar fashion.
*
* May take several minutes of run time when there are many files and
* subfolders in the folder, and may even time out.
*
* @param {Folder} folder The top-level folder to recurse.
* @param {String[]} folderIds An array where the IDs of visited subfolders are concatenated to avoid circular recursion.
* @param {String[][]} result A 2D array where the results are concatenated.
* @param {String} indent A text string to prepend to subfolder names to show nesting nevel.
* @return {String[][]} A list of names, types, URLs and sizes of files by subfolder.
*/
function getFileInfoBySubFolder_(folder, folderIds, indent, result) {
  // version 1.1, written by --Hyde, 15 July 2020
  //  - add recursion
  //  - add file size and file type info
  //  - add indent to show nesting level
  //  - see https://support.google.com/docs/thread/58260211
  // version 1.0, written by --Hyde, 16 June 2020
  //  - initial version
  //  - see https://support.google.com/docs/thread/53541982
  const folderId = folder.getId();
  if (!folderIds) {
    folderIds = [];
  }
  if (folderIds.indexOf(folderId) !== -1) { // avoid circular recursion
    return result;
  }
  folderIds.push(folderId);
  if (result === undefined) {
    result = [];
  }
  const folderInfo = getFolderInfo_(folder, indent || '');
  const fileInfo = getFolderFileInfo_(folder) || [];
  if (fileInfo.length) { // sum file sizes in last column of fileInfo and store in folderInfo
    let lastColumn = fileInfo[0].length - 1;
    folderInfo[lastColumn] = tableColumnSum_(fileInfo, lastColumn);
  }
  result.push(folderInfo);
  result = result.concat(fileInfo);
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    result = getFileInfoBySubFolder_(subfolder, folderIds, '    ' + (indent || '↳'), result);
  }
  return result;
}

/**
* Gets the name, URL and size of a folder.
*
* @param {Folder} folder The folder.
* @param {String} indent A text string to prepend to subfolder names to show nesting nevel.
* @return {String[]} The name of the folder, a null value, a null value, the URL and size of the folder.
*/
function getFolderInfo_(folder, indent) {
  // version 1.1, written by --Hyde, 15 July 2020
  //  - add folder size info
  //  - use folder instead of folderId
  //  - see https://support.google.com/docs/thread/58260211
  // version 1.0, written by --Hyde, 16 June 2020
  //  - initial version
  //  - see https://support.google.com/docs/thread/53541982
  return [indent + folder.getName(), null, null, folder.getUrl(), folder.getSize() || null];
}

/**
* Gets the names, URLs and file sizes of files in a folder.
*
* @param {String} folder The folder where to get filenames, URLs and file sizes.
* @return {String[][]} The filenames, URLs and sizes, preceded by a null, or null if there are no files in the folder.
*/
function getFolderFileInfo_(folder) {
  // version 1.2, written by --Hyde, 26 May 2021
  //  - sort result by filename
  // version 1.1, written by --Hyde, 15 July 2020
  //  - add file size and file type info
  //  - use folder instead of folderId
  //  - see https://support.google.com/docs/thread/58260211
  // version 1.0, written by --Hyde, 16 June 2020
  //  - initial version
  //  - see https://support.google.com/docs/thread/53541982
  const files = folder.getFiles();
  let result = [];
  while (files.hasNext()) {
    let file = files.next();
    result.push([null, file.getName(), file.getMimeType(), file.getUrl(), file.getSize()]);
  }
  return result.length ? result.sort((row1, row2) => {
    const file1 = row1[1].toLowerCase();
    const file2 = row2[1].toLowerCase();
    switch (true) {
      case file1 < file2:
        return -1;
      case file1 > file2:
        return 1;
      default:
        return 0;
    }
  }) : null;
}

/**
* Adds together numbers in a column in a 2D array.
*
* Coerses numbers and text strings that look like numbers.
* Ignores other kinds of values.
*
* @param {Object[][]} rows The 2D array where to sum numbers in a column.
* @param {Number} column The column to sum.
* @return {Number} The sum of numbers in column.
*/
function tableColumnSum_(rows, column) {
  // version 1.0, written by --Hyde, 15 July 2020
  //  - initial version
  return rows
    .map(function (row) {
      return row[column];
    })
    .reduce(function (sum, value) {
      if (Number(value)) {
        return sum + Number(value);
      }
      return sum;
    }, 0);
}

/**
* Shows a dialog box that lets the user enter a file name or ID.
* Then finds the file in Google Drive.
*
* @return {Object} Null if the user clicked Cancel. Otherwise a result object with these fields:
*                  {File} file The file the user selected, or null if the file could not be found.
*                  {String} fileId The file's ID number.
*                  {String} fileType The file's MIME type.
*                  {String} error An error message, or null if there are no errors.
*/
function chooseFileOrFolder_() {
  // version 1.0, written by --Hyde, 15 July 2020
  //  - initial version
  const result = {
    file: null,
    fileId: null,
    fileType: null,
    error: null,
  };
  const ui = SpreadsheetApp.getUi();
  let response;
  const userResponse = ui.prompt(
    'List details of files in a folder',
    'This utility gets the names, types, URLs and sizes of files in a folder and its subfolders, and ' +
    'inserts the list in the spreadsheet starting at the active cell.\n\n' +
    'Existing values will be overwritten in the range that starts at the active cell and extends down and to the right.\n\n' +
    'Enter the name or ID of the folder (or a single file) to list:\n\n',
    ui.ButtonSet.OK_CANCEL);
  switch (userResponse.getSelectedButton()) {
    case ui.Button.OK:
      response = userResponse.getResponseText();
      if (response === '') {
        return null;
      }
      break;
    case ui.Button.CANCEL:
    case ui.Button.CLOSE:
    default:
      return null;
  }
  let fileId;
  try {
    fileId = DriveApp.getFileById(response).getId();
  } catch (error) {
    try {
      fileId = DriveApp.getFilesByName(response).next().getId();
    } catch (error) {
      try {
        fileId = DriveApp.getFoldersByName(response).next().getId();
      } catch (error) {
        result.error = 'Cannot find a file or folder by the name or ID "' + response + '".';
        return result;
      }
    }
  }
  result.file = DriveApp.getFileById(fileId);
  result.fileId = fileId;
  result.fileType = result.file.getMimeType();
  return result;
}

/**
* Shows error.message in a pop-up and throws the error.
*
* @param {Error} error The error to show and throw.
*/
function showAndThrow_(error) {
  // version 1.0, written by --Hyde, 16 April 2020
  //  - initial version
  var stackCodeLines = String(error.stack).match(/\d+:/);
  if (stackCodeLines) {
    var codeLine = stackCodeLines.join(', ').slice(0, -1);
  } else {
    codeLine = error.stack;
  }
  showMessage_(error.message + ' Code line: ' + codeLine, 30);
  throw error;
}
/**
* Shows a message in a pop-up.
*
* @param {String} message The message to show.
* @param {Number} timeoutSeconds Optional. The number of seconds before the message goes away. Defaults to 5.
*/
function showMessage_(message, timeoutSeconds) {
  // version 1.0, written by --Hyde, 16 April 2020
  //  - initial version
  SpreadsheetApp.getActive().toast(message, 'File sizes script', timeoutSeconds || 5);
}
