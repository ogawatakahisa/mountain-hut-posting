// spreadsheetUtils.js

// 既存の関数内容の全てを書き出します。

function getSpreadsheetByName(folderPath, spreadsheetName) {
  const folderNames = folderPath.split('/');
  let folder = null;
  try {
    const rootFolders = DriveApp.getFoldersByName(folderNames[0]);
    if (rootFolders.hasNext()) {
      folder = rootFolders.next();
    } else {
      console.error(`Root folder "${folderNames[0]}" not found.`);
      return null;
    }
    for (let i = 1; i < folderNames.length; i++) {
      const folders = folder.getFoldersByName(folderNames[i]);
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        console.error(`Folder "${folderNames[i]}" not found in "${folderNames.slice(0, i).join('/')}".`);
        return null;
      }
    }
    const files = folder.getFilesByName(spreadsheetName);
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === "application/vnd.google-apps.spreadsheet") {
        return SpreadsheetApp.openById(file.getId());
      }
    }
    console.error(`Spreadsheet "${spreadsheetName}" not found in folder "${folderPath}".`);
    return null;
  } catch (error) {
    console.error(`Error while searching for spreadsheet: ${error.message}`);
    return null;
  }
}

function findColumn(sheet, columnName) {
  let i = 7;
  let range = sheet.getRange(i, 1, 1, sheet.getLastColumn());
  let values = range.getValues()[0];
  return values.indexOf(columnName);
}

function findFirstEmptyRow(sheet, columnIndex) {
  let lastRow = sheet.getLastRow();
  for (let row = 8; row <= lastRow; row++) {
    let cell = sheet.getRange(row, columnIndex + 1);
    if (!cell.getValue()) {
      return row;
    }
  }
  return lastRow + 1;
}
