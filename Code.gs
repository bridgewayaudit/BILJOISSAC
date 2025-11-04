// Apps Script: Code.gs
// NOTE: Set SHEET_ID and ATTACHMENTS_FOLDER_ID below to match your environment.

var SHEET_ID = 'YOUR_SHEET_ID_HERE'; // e.g. '1AbCdEf...'
var SHEET_NAME = 'Sheet1'; // change to your sheet name
var ATTACHMENTS_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE'; // change to target Drive folder

/**
 * Entrypoint for web app. Supports actions: getData, addRow, updateRow
 */
function doPost(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || (e.postData && e.postData.type === 'application/json' && JSON.parse(e.postData.contents).action) || '';
    var dataStr = params.data || (e.postData && e.postData.type === 'application/json' && JSON.parse(e.postData.contents).data) || '';
    var payload = {};
    try { payload = dataStr && typeof dataStr === 'string' ? JSON.parse(dataStr) : (dataStr || {}); } catch(err) { payload = dataStr || {}; }

    if (action === 'addRow') {
      var row = addRowToSheet(payload);
      return ContentService.createTextOutput(JSON.stringify({ success: true, row: row })).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'updateRow') {
      var updated = updateRowInSheet(payload);
      return ContentService.createTextOutput(JSON.stringify({ success: true, row: updated })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var params = e.parameter || {};
  var action = params.action || '';
  if (action === 'getData') {
    var data = getAllRows();
    return ContentService.createTextOutput(JSON.stringify({ data: data })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Add a row object (data) to the sheet. Handles data.Attachments (JSON array) by saving files to Drive and
 * storing AttachmentUrls as a JSON array string in the 'AttachmentUrls' column.
 */
function addRowToSheet(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];

  // Process attachments if provided
  if (data.Attachments) {
    try {
      var attachments = typeof data.Attachments === 'string' ? JSON.parse(data.Attachments) : data.Attachments;
      var saved = saveAttachmentsToDrive(attachments);
      // Save URLs array and full metadata array
      data.AttachmentUrls = JSON.stringify(saved.map(function(x){ return x.url; }));
      data.AttachmentsMeta = JSON.stringify(saved);
    } catch (e) {
      // log and continue
      Logger.log('Attachment processing failed: ' + e);
    }
  }

  // Prepare headers
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  // If sheet is empty, write headers from keys of data
  if (!headers || headers.length === 0 || headers.every(function(h){ return h === '';})) {
    headers = Object.keys(data);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  }

  // Map data to row according to headers
  var row = headers.map(function(h){ return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(row);

  // Return the inserted row object (as stored)
  var numRows = sheet.getLastRow();
  var values = sheet.getRange(numRows,1,1,headers.length).getValues()[0];
  var out = {};
  headers.forEach(function(h,i){ out[h] = values[i]; });
  return out;
}

/**
 * Update an existing row by matching SlNo field. It merges incoming fields and also processes Attachments.
 * Expects payload to contain SlNo.
 */
function updateRowInSheet(data) {
  if (!data.SlNo) throw new Error('SlNo is required to update a row');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];

  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var rows = sheet.getRange(2,1,Math.max(0,sheet.getLastRow()-1), headers.length).getValues();

  var targetRowIndex = -1;
  for (var i=0;i<rows.length;i++){
    var rowObj = {};
    headers.forEach(function(h,idx){ rowObj[h] = rows[i][idx]; });
    if (String(rowObj.SlNo || '') === String(data.SlNo)) { targetRowIndex = i+2; break; }
  }
  if (targetRowIndex === -1) throw new Error('Row with SlNo not found');

  // Process attachments
  if (data.Attachments) {
    try {
      var attachments = typeof data.Attachments === 'string' ? JSON.parse(data.Attachments) : data.Attachments;
      var saved = saveAttachmentsToDrive(attachments);
      // If existing AttachmentUrls column has previous URLs, merge
      var existing = sheet.getRange(targetRowIndex, headers.indexOf('AttachmentUrls')+1).getValue();
      var existingArr = [];
      try { existingArr = existing ? JSON.parse(existing) : []; } catch(e){ existingArr = existing ? [existing] : []; }
      var newUrls = saved.map(function(x){ return x.url; });
      var merged = existingArr.concat(newUrls);
      data.AttachmentUrls = JSON.stringify(merged);
      // store metadata as well
      var existingMeta = sheet.getRange(targetRowIndex, headers.indexOf('AttachmentsMeta')+1).getValue();
      var existingMetaArr = [];
      try { existingMetaArr = existingMeta ? JSON.parse(existingMeta) : []; } catch(e) { existingMetaArr = existingMeta ? [existingMeta] : []; }
      data.AttachmentsMeta = JSON.stringify(existingMetaArr.concat(saved));
    } catch (e) { Logger.log('Attachment processing failed on update: ' + e); }
  }

  // Merge new data into the row
  var finalRow = headers.map(function(h){ return data[h] !== undefined ? data[h] : sheet.getRange(targetRowIndex, headers.indexOf(h)+1).getValue(); });
  sheet.getRange(targetRowIndex,1,1,headers.length).setValues([finalRow]);

  // Return updated row
  var values = sheet.getRange(targetRowIndex,1,1,headers.length).getValues()[0];
  var out = {};
  headers.forEach(function(h,i){ out[h] = values[i]; });
  return out;
}

/**
 * Save attachments array to Drive folder and return array of saved file metadata: {name,mime,url,id,size}
 * attachments: array of {name,mime,base64} or {name,url}
 */
function saveAttachmentsToDrive(attachments) {
  if (!attachments || !attachments.length) return [];
  var folder;
  try { folder = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID); } catch(e) { folder = DriveApp.getRootFolder(); }
  var saved = [];
  attachments.forEach(function(att){
    try {
      if (typeof att === 'string') {
        // treat as URL
        saved.push({ name: att.split('/').pop(), url: att });
        return;
      }
      if (att.url) {
        saved.push({ name: att.name || (att.url.split('/').pop()), mime: att.mime || 'application/octet-stream', url: att.url, id: att.id || '' });
        return;
      }
      if (att.base64) {
        var bytes = Utilities.base64Decode(att.base64);
        var blob = Utilities.newBlob(bytes, att.mime || 'application/octet-stream', att.name || ('file-' + new Date().getTime()));
        var file = folder.createFile(blob);
        // make file accessible by anyone with link (optional; adjust as needed)
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) { Logger.log('setSharing failed: ' + e); }
        var url = file.getUrl();
        saved.push({ name: file.getName(), mime: blob.getContentType(), url: url, id: file.getId(), size: blob.getBytes().length });
      }
    } catch (e) {
      Logger.log('saveAttachmentsToDrive error for item: ' + JSON.stringify(att) + ' error: ' + e);
    }
  });
  return saved;
}

/**
 * Read all rows as array of objects
 */
function getAllRows() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[0];
  var range = sheet.getDataRange();
  var vals = range.getValues();
  if (vals.length < 2) return [];
  var headers = vals[0];
  var out = [];
  for (var r=1;r<vals.length;r++){
    var obj = {};
    for (var c=0;c<headers.length;c++){
      obj[headers[c]] = vals[r][c];
    }
    out.push(obj);
  }
  return out;
}

/**
 * Utility: ensure columns exist - optionally call this to add AttachmentUrls and AttachmentsMeta headers if missing
 */
function ensureAttachmentColumns() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var headers = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getValues()[0];
  var toAdd = [];
  if (headers.indexOf('AttachmentUrls') === -1) toAdd.push('AttachmentUrls');
  if (headers.indexOf('AttachmentsMeta') === -1) toAdd.push('AttachmentsMeta');
  if (toAdd.length) {
    sheet.getRange(1, headers.length+1, 1, toAdd.length).setValues([toAdd]);
  }
}

// End of Code.gs