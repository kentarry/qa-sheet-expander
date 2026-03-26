/*
  品檢展表工具 — Drive 讀取 API v2（快速版）
  
  部署步驟：
  1. 前往 https://script.google.com
  2. 建立新專案，貼上此程式碼
  3. 部署 → 新增部署 → 網頁應用程式
  4. 執行身份：我 / 存取權：所有人
  5. 複製網址貼到 src/config.js
  
  API：
  ?action=list&folderId=xxx  → 列出子資料夾和 Excel 檔案
  ?action=download&fileId=xxx → 下載檔案回傳 base64
*/

function doGet(e) {
  var action = e.parameter.action || 'list';
  var result;
  try {
    if (action === 'list') {
      result = listFolder(e.parameter.folderId);
    } else if (action === 'download') {
      result = downloadFile(e.parameter.fileId);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function listFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var folders = [], files = [];
  var fi = folder.getFolders();
  while (fi.hasNext()) {
    var f = fi.next();
    folders.push({ id: f.getId(), name: f.getName() });
  }
  folders.sort(function(a, b) { return a.name.localeCompare(b.name); });
  var fls = folder.getFiles();
  while (fls.hasNext()) {
    var file = fls.next();
    var mime = file.getMimeType();
    var name = file.getName();
    // 擴大 MIME 過濾：spreadsheet, excel, ms-excel, openxmlformats, 或副檔名 .xls/.xlsx
    var isExcel = mime.indexOf('spreadsheet') >= 0 || mime.indexOf('excel') >= 0 || mime.indexOf('ms-excel') >= 0 || mime.indexOf('openxmlformats') >= 0 || /\.(xlsx?|xls)$/i.test(name);
    if (isExcel) {
      files.push({ id: file.getId(), name: name, updated: file.getLastUpdated().toISOString().slice(0, 10) });
    }
  }
  files.sort(function(a, b) { return b.updated.localeCompare(a.updated); });
  return { folderName: folder.getName(), folders: folders, files: files };
}

function downloadFile(fileId) {
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var mime = file.getMimeType();
  if (mime === 'application/vnd.google-apps.spreadsheet') {
    var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
    var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } });
    blob = response.getBlob();
  }
  return { fileId: fileId, fileName: file.getName(), base64: Utilities.base64Encode(blob.getBytes()) };
}
