/**
 * Google Apps Script for managing file uploads from Streamlit to Google Drive
 * and updating the associated Google Sheet.
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var fileName = data.fileName;
    var fileData = data.fileData; // base64 string
    var mimeType = data.mimeType;
    var folderName = data.folderName; // Applicant Name
    var sheetId = data.sheetId;
    var sheetName = data.sheetName;
    var rowIdx = data.rowIdx; // 1-indexed row number from Streamlit
    
    // 1. Get or Create the main "ระบบจัดการใบอนุญาต" folder
    var mainFolderName = "ระบบจัดการใบอนุญาต";
    var mainFolder = getOrCreateFolder(DriveApp, mainFolderName);
    
    // 2. Get or Create "Attachments" folder inside main folder
    var attachmentFolder = getOrCreateFolder(mainFolder, "Attachments");
    
    // 3. Get or Create subfolder for the applicant
    var personFolder = getOrCreateFolder(attachmentFolder, folderName);
    
    // 4. Save the file
    var bytes = Utilities.base64Decode(fileData);
    var blob = Utilities.newBlob(bytes, mimeType, fileName);
    var file = personFolder.createFile(blob);
    var fileUrl = file.getUrl();
    
    // 5. Update the Google Sheet
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.getSheets()[0]; // Fallback to first sheet
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colIdx = -1;
    
    // Find "ไฟล์แนบ" column
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].toString().trim() === "ไฟล์แนบ") {
        colIdx = i + 1;
        break;
      }
    }
    
    // If column doesn't exist, create it
    if (colIdx === -1) {
      colIdx = headers.length + 1;
      sheet.getRange(1, colIdx).setValue("ไฟล์แนบ");
    }
    
    // Set the file URL in the sheet
    sheet.getRange(rowIdx, colIdx).setValue(fileUrl);
    
    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "url": fileUrl,
      "fileName": fileName
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper function to find or create a folder
 */
function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parent.createFolder(name);
  }
}
