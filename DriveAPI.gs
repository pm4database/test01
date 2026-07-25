/**
 * =========================================================================
 * DriveAPI.gs — GAS Web App สำหรับ Google Drive Operations
 * =========================================================================
 * Deploy เป็น Web App:
 *   Execute as: Me (owner)
 *   Who has access: Anyone
 *
 * Actions:
 *   uploadTemplateImage  → อัปโหลดรูปพื้นหลัง template
 *   uploadElementImage   → อัปโหลดรูป element
 *   saveCertificateImage → บันทึกเกียรติบัตรที่สร้างแล้ว
 *   saveZipFile          → บันทึก ZIP ไฟล์
 *   getImageBase64       → ดึงรูปจาก Drive เป็น base64
 * =========================================================================
 */


// ═══════════════════════════════════════════════════════════════════════
// WEB APP ENDPOINTS
// ═══════════════════════════════════════════════════════════════════════

/**
 * doPost(e) — รับ request จาก frontend
 * Body: JSON { action: 'uploadTemplateImage', base64Data: '...', filename: '...' }
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || '';

    switch (action) {
      case 'uploadTemplateImage':
        return _respond(_uploadTemplateImage(body));

      case 'uploadElementImage':
        return _respond(_uploadElementImage(body));

      case 'saveCertificateImage':
        return _respond(_saveCertificateImage(body));

      case 'saveZipFile':
        return _respond(_saveZipFile(body));

      case 'getImageBase64':
        return _respond(_getImageBase64(body));

      case 'healthCheck':
        return _respond(_healthCheck());

      default:
        return _respond({ status: false, message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return _respond({ status: false, message: 'Server Error: ' + err.message });
  }
}


/**
 * doGet(e) — Health Check / CORS preflight
 */
function doGet(e) {
  return _respond(_healthCheck());
}


// ═══════════════════════════════════════════════════════════════════════
// ACTION HANDLERS
// ═══════════════════════════════════════════════════════════════════════

/**
 * อัปโหลดรูปพื้นหลัง Template → Templates folder
 */
function _uploadTemplateImage(body) {
  _validateUpload(body);

  var folder = _getFolder('DRIVE_TEMPLATE_FOLDER');
  var blob = _base64ToBlob(body.base64Data, body.filename);

  // ลบไฟล์ซ้ำชื่อเดิม (ถ้ามี)
  _removeDuplicateFile(folder, body.filename);

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    status: true,
    file_id: file.getId(),
    fileId: file.getId(),
    url: file.getUrl(),
    name: file.getName(),
    message: 'อัปโหลดพื้นหลังสำเร็จ'
  };
}


/**
 * อัปโหลดรูป Element → Templates folder
 */
function _uploadElementImage(body) {
  _validateUpload(body);

  var folder = _getFolder('DRIVE_TEMPLATE_FOLDER');
  var blob = _base64ToBlob(body.base64Data, body.filename);

  _removeDuplicateFile(folder, body.filename);

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    status: true,
    file_id: file.getId(),
    fileId: file.getId(),
    url: file.getUrl(),
    name: file.getName(),
    message: 'อัปโหลดรูปภาพสำเร็จ'
  };
}


/**
 * บันทึกเกียรติบัตรที่ generate แล้ว → Generated folder
 * Body: { base64Data, filename, rowIndex, templateName }
 */
function _saveCertificateImage(body) {
  _validateUpload(body);

  var parentFolder = _getFolder('DRIVE_GENERATED_FOLDER');

  // สร้าง sub-folder ตามชื่อ template (ถ้ามี)
  var folder = parentFolder;
  if (body.templateName) {
    var subFolders = parentFolder.getFoldersByName(body.templateName);
    if (subFolders.hasNext()) {
      folder = subFolders.next();
    } else {
      folder = parentFolder.createFolder(body.templateName);
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
  }

  var blob = _base64ToBlob(body.base64Data, body.filename);

  _removeDuplicateFile(folder, body.filename);

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    status: true,
    file_id: file.getId(),
    fileId: file.getId(),
    url: file.getUrl(),
    driveUrl: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    name: file.getName(),
    rowIndex: body.rowIndex,
    message: 'บันทึกเกียรติบัตรสำเร็จ'
  };
}


/**
 * บันทึก ZIP ไฟล์ → ZIP folder
 * Body: { base64Data, filename }
 */
function _saveZipFile(body) {
  _validateUpload(body);

  var folder = _getFolder('DRIVE_ZIP_FOLDER');
  var blob = Utilities.newBlob(
    Utilities.base64Decode(body.base64Data),
    'application/zip',
    body.filename
  );

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    status: true,
    file_id: file.getId(),
    fileId: file.getId(),
    url: file.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    name: file.getName(),
    message: 'บันทึก ZIP สำเร็จ'
  };
}


/**
 * ดึงรูปจาก Drive เป็น base64
 * Body: { fileId }
 */
function _getImageBase64(body) {
  if (!body.fileId) {
    return { status: false, message: 'ไม่ได้ระบุ fileId' };
  }

  try {
    var file = DriveApp.getFileById(body.fileId);
    var blob = file.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    var mimeType = blob.getContentType() || 'image/png';

    return {
      status: true,
      base64: base64,
      dataUrl: 'data:' + mimeType + ';base64,' + base64,
      mimeType: mimeType,
      name: file.getName(),
      size: blob.getBytes().length
    };
  } catch (err) {
    return { status: false, message: 'ไม่พบไฟล์: ' + err.message };
  }
}


/**
 * Health Check
 */
function _healthCheck() {
  var props = PropertiesService.getScriptProperties();
  return {
    status: true,
    service: 'DriveAPI',
    version: '2.0.0',
    timestamp: new Date().toISOString(),
    folders: {
      template: props.getProperty('DRIVE_TEMPLATE_FOLDER') ? '✅' : '❌',
      generated: props.getProperty('DRIVE_GENERATED_FOLDER') ? '✅' : '❌',
      zip: props.getProperty('DRIVE_ZIP_FOLDER') ? '✅' : '❌',
      temp: props.getProperty('DRIVE_TEMP_FOLDER') ? '✅' : '❌'
    }
  };
}


// ═══════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════

/**
 * ดึง Folder จาก Script Properties
 * ใช้ CacheService เพื่อลดการอ่าน Properties ซ้ำ
 */
function _getFolder(propKey) {
  var cache = CacheService.getScriptCache();
  var folderId = cache.get(propKey);

  if (!folderId) {
    folderId = PropertiesService.getScriptProperties().getProperty(propKey);
    if (folderId) {
      cache.put(propKey, folderId, 3600); // cache 1 ชม.
    }
  }

  if (!folderId) {
    throw new Error('ไม่พบ Folder ID สำหรับ ' + propKey + ' — กรุณารัน setupDriveFoldersOnly() ก่อน');
  }

  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    // Folder ถูกลบ → clear cache แล้ว error
    cache.remove(propKey);
    throw new Error('Folder ไม่พบ (อาจถูกลบ): ' + folderId);
  }
}


/**
 * แปลง base64 → Blob
 */
function _base64ToBlob(base64Data, filename) {
  // กำจัด data URL prefix ถ้ามี
  if (base64Data.indexOf(',') !== -1) {
    base64Data = base64Data.split(',')[1];
  }

  var decoded = Utilities.base64Decode(base64Data);

  // ตรวจ MIME type จากนามสกุลไฟล์
  var ext = (filename || '').split('.').pop().toLowerCase();
  var mimeTypes = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'webp': 'image/webp',
    'svg': 'image/svg+xml',
    'zip': 'application/zip'
  };
  var mimeType = mimeTypes[ext] || 'image/png';

  return Utilities.newBlob(decoded, mimeType, filename || 'upload.png');
}


/**
 * ตรวจสอบ upload request
 */
function _validateUpload(body) {
  if (!body.base64Data) {
    throw new Error('ไม่มีข้อมูลไฟล์ (base64Data)');
  }

  // จำกัดขนาดไฟล์ 10MB (base64 ≈ 1.37x ของ binary)
  var estimatedSize = body.base64Data.length * 0.75;
  var maxSize = 10 * 1024 * 1024; // 10MB
  if (estimatedSize > maxSize) {
    throw new Error('ไฟล์ใหญ่เกินไป (' + Math.round(estimatedSize / 1024 / 1024) + 'MB) — จำกัด 10MB');
  }

  if (!body.filename) {
    body.filename = 'upload_' + Date.now() + '.png';
  }
}


/**
 * ลบไฟล์ซ้ำชื่อเดิมใน folder (ถ้ามี)
 */
function _removeDuplicateFile(folder, filename) {
  try {
    var files = folder.getFilesByName(filename);
    while (files.hasNext()) {
      var f = files.next();
      f.setTrashed(true);
    }
  } catch (e) {
    // ignore — ถ้าลบไม่ได้ก็ไม่เป็นไร
  }
}


/**
 * สร้าง JSON Response
 */
function _respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
