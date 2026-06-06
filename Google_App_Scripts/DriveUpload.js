// =============================================
// DRIVEUPLOAD.gs - Save images to Google Drive
// =============================================

// --------------------------------------------
// UPLOAD: Receives base64 data from the form,
// saves to the configured Drive folder,
// returns the public file URL
// --------------------------------------------
function uploadImageToDrive(base64Data, fileName, mimeType) {
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
    if (!folderId) throw new Error('DRIVE_FOLDER_ID is not set in Script Properties.');

    var folder  = DriveApp.getFolderById(folderId);
    var decoded = Utilities.base64Decode(base64Data);
    var blob    = Utilities.newBlob(decoded, mimeType, fileName);
    var file    = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId  = file.getId();
    var fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

    // Return a plain object - no prototype methods, safe for google.script.run
    return {
      success:  true,
      fileId:   fileId,
      fileUrl:  fileUrl,
      fileName: file.getName()
    };

  } catch (err) {
    Logger.log('uploadImageToDrive error: ' + err.message);
    // Must return a plain serializable object - never throw across google.script.run
    return {
      success: false,
      error:   err.message
    };
  }
}

// --------------------------------------------
// LIST: Get all images in the CMS Drive folder
// Useful for a future media picker
// --------------------------------------------
function getDriveFolderImages() {
  try {
    var folderId = CONFIG.DRIVE_FOLDER_ID;
    if (!folderId) throw new Error('DRIVE_FOLDER_ID is not set in Script Properties.');

    var folder  = DriveApp.getFolderById(folderId);
    var files   = folder.getFiles();
    var results = [];

    while (files.hasNext()) {
      var file     = files.next();
      var mimeType = file.getMimeType();

      // Only return image files
      if (mimeType.indexOf('image/') === 0) {
        results.push({
          fileId:   file.getId(),
          fileName: file.getName(),
          fileUrl:  'https://drive.google.com/uc?export=view&id=' + file.getId(),
          mimeType: mimeType,
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