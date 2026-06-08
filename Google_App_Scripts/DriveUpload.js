// =============================================
// DRIVEUPLOAD.gs - Save images to Google Drive
// =============================================

// --------------------------------------------
// AUTO-FOLDER: Get the configured Drive folder,
// or create one automatically if DRIVE_FOLDER_ID
// is not set. Saves the new ID to Script Properties
// so it persists for future uploads.
// --------------------------------------------
function getOrCreateCMSFolder() {
  var props    = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('DRIVE_FOLDER_ID');

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      Logger.log('Stored DRIVE_FOLDER_ID invalid, will create new folder: ' + e.message);
    }
  }

  var folderName = 'JPAR CMS Media';
  var existing   = DriveApp.getFoldersByName(folderName);
  var folder     = existing.hasNext() ? existing.next() : DriveApp.createFolder(folderName);

  // Make files in the folder viewable by anyone with the link
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Persist so next call uses the same folder
  props.setProperty('DRIVE_FOLDER_ID', folder.getId());
  Logger.log('CMS Drive folder ready: ' + folder.getId());

  return folder;
}

// --------------------------------------------
// UPLOAD: Receives base64 data from the form,
// saves to the CMS Drive folder,
// returns the public file URL
// --------------------------------------------
function uploadImageToDrive(base64Data, fileName, mimeType) {
  try {
    var folder  = getOrCreateCMSFolder();
    var decoded = Utilities.base64Decode(base64Data);
    var blob    = Utilities.newBlob(decoded, mimeType, fileName);
    var file    = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId  = file.getId();
    var fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    var warning = null;

    // setSharing may be blocked by Google Workspace org policies.
    // Wrap separately so the file URL is still returned even if sharing fails.
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      Logger.log('setSharing blocked (org restriction?): ' + sharingErr.message);
      warning = 'File saved but public sharing is restricted by your Google Workspace admin. The image may not display on the public site.';
    }

    return {
      success:  true,
      fileId:   fileId,
      fileUrl:  fileUrl,
      fileName: file.getName(),
      warning:  warning
    };

  } catch (err) {
    Logger.log('uploadImageToDrive error: ' + err.message);
    return {
      success: false,
      error:   err.message
    };
  }
}

// --------------------------------------------
// LIST: Get all images in the CMS Drive folder
// --------------------------------------------
function getDriveFolderImages() {
  try {
    var folder  = getOrCreateCMSFolder();
    var files   = folder.getFiles();
    var results = [];

    while (files.hasNext()) {
      var file = files.next();
      if (file.getMimeType().indexOf('image/') === 0) {
        results.push({
          fileId:   file.getId(),
          fileName: file.getName(),
          fileUrl:  'https://drive.google.com/uc?export=view&id=' + file.getId(),
          mimeType: file.getMimeType(),
          created:  file.getDateCreated().toISOString()
        });
      }
    }
    return results;

  } catch (err) {
    Logger.log('getDriveFolderImages error: ' + err.message);
    return [];
  }
}
