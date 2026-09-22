/**
 * The one dedicated Drive folder for this workbook: receipt images and
 * generated Tax Export PDFs both live here. The folder ID is cached in
 * Script Properties so we only create it once.
 */

function getOrCreateReceiptsFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(PROPS_RECEIPTS_FOLDER_ID);

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // Stored ID no longer resolves (folder deleted/trashed) - fall through and recreate.
    }
  }

  var folder = findOrCreateFolderNextToSpreadsheet();
  props.setProperty(PROPS_RECEIPTS_FOLDER_ID, folder.getId());
  return folder;
}

/**
 * Prefers creating the receipts folder alongside the spreadsheet (in the
 * same parent Drive folder) so it's easy to find; falls back to Drive root
 * if the spreadsheet has no accessible parent (e.g. it's in "Shared with me").
 */
function findOrCreateFolderNextToSpreadsheet() {
  var ssFile = DriveApp.getFileById(getSs().getId());
  var parents = ssFile.getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  var existing = parent.getFoldersByName(RECEIPTS_FOLDER_NAME);
  if (existing.hasNext()) return existing.next();

  return parent.createFolder(RECEIPTS_FOLDER_NAME);
}

function getReceiptsFolderUrl() {
  return getOrCreateReceiptsFolder().getUrl();
}

/**
 * Shows a small dialog with a clickable link to the receipts folder
 * (Apps Script server code can't open a browser tab directly).
 */
function showReceiptsFolderLink() {
  var url = getReceiptsFolderUrl();
  var html = HtmlService.createHtmlOutput(
    '<p style="font-family:Arial,sans-serif;font-size:13px;">' +
    'Receipts folder:</p>' +
    '<p><a href="' + url + '" target="_blank">' + url + '</a></p>'
  ).setWidth(420).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Budget Tracker Receipts');
}

/**
 * Uploads a receipt file (from the sidebar's <input type="file">) to the
 * receipts folder, makes it viewable via link, and returns its URL.
 *
 * @param {Object} fileData {name, mimeType, base64} - base64 has no
 *   "data:...;base64," prefix (stripped client-side).
 * @param {Object} meta {date, scope, category, amount} used to build a
 *   human-readable, audit-friendly file name.
 * @return {string} The shareable view URL.
 */
function uploadReceiptFile(fileData, meta) {
  if (!fileData || !fileData.base64) {
    throw new Error('No receipt file data received.');
  }

  var bytes = Utilities.base64Decode(fileData.base64);
  if (bytes.length > MAX_RECEIPT_FILE_BYTES) {
    throw new Error('Receipt file is too large (max 10 MB).');
  }

  var mimeType = fileData.mimeType || 'application/octet-stream';
  var extension = extensionFromFileName(fileData.name) || extensionFromMimeType(mimeType);
  var fileName = buildReceiptFileName(meta, extension);

  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var folder = getOrCreateReceiptsFolder();
  var file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Sharing can be restricted by a Workspace admin policy - the file
    // still exists and is usable by the owner even if this fails.
  }

  return file.getUrl();
}

function buildReceiptFileName(meta, extension) {
  var dateStr = meta && meta.date ? meta.date : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var scope = sanitizeForFileName(meta && meta.scope);
  var category = sanitizeForFileName(meta && meta.category) || UNCATEGORISED_LABEL;
  var amount = meta && meta.amount ? Number(meta.amount).toFixed(2) : '0.00';
  var parts = [dateStr, scope, category, amount].filter(function (p) { return p; });
  return parts.join('_') + (extension ? '.' + extension : '');
}

function sanitizeForFileName(str) {
  return (str || '').replace(/[^a-zA-Z0-9 &-]/g, '').trim().replace(/\s+/g, '-');
}

function extensionFromFileName(name) {
  if (!name) return '';
  var m = /\.([a-zA-Z0-9]+)$/.exec(name);
  return m ? m[1].toLowerCase() : '';
}

function extensionFromMimeType(mimeType) {
  var map = {
    'image/jpeg': 'jpg',
    'image/png': 'png',
    'image/heic': 'heic',
    'image/webp': 'webp',
    'application/pdf': 'pdf'
  };
  return map[mimeType] || '';
}
