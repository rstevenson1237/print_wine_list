// utils.gs

/**
 * @fileoverview Utility functions used across the application.
 */

/**
 * Escapes HTML special characters in a string.
 * @param {string} text The text to escape.
 * @returns {string} The escaped text.
 */
function escapeHtml(text) {
  if (!text) return '';
  var map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  return text.toString().replace(/[&<>"']/g, function(m) { return map[m]; });
}

/**
 * Creates a URL-friendly slug from a string for use in HTML IDs.
 * @param {string} text The text to convert.
 * @returns {string} The slugified text.
 */
function createSlug(text) {
  if (!text) return '';
  return text.toString().toLowerCase()
    .replace(/\s+/g, '-')
    .replace(/[^\w\-]+/g, '')
    .replace(/\-\-+/g, '-')
    .replace(/^-+/, '')
    .replace(/-+$/, '');
}

/**
 * Finds a file in the spreadsheet's parent Drive folder (or general Drive)
 * and returns it as a Base64 data URI.
 *
 * An explicit mimeType can be provided to override the content type returned
 * by Drive (necessary for font files, which Drive often reports as
 * application/octet-stream).
 *
 * @param {string} fileName The name of the file to find.
 * @param {string} [mimeType] Optional MIME type override.
 * @returns {string|null} A data URI string, or null if not found.
 */
function getFileAsDataUri(fileName, mimeType) {
  try {
    var file = null;

    // Search spreadsheet's parent folder first
    var ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var parents = ssFile.getParents();
    if (parents.hasNext()) {
      var results = parents.next().getFilesByName(fileName);
      if (results.hasNext()) {
        file = results.next();
      }
    }

    // Fall back to general Drive search
    if (!file) {
      var results2 = DriveApp.getFilesByName(fileName);
      if (results2.hasNext()) {
        file = results2.next();
      }
    }

    if (!file) {
      Logger.log('File not found: ' + fileName);
      return null;
    }

    var blob = file.getBlob();
    var resolvedMime = mimeType || blob.getContentType();
    var base64Data = Utilities.base64Encode(blob.getBytes());
    return 'data:' + resolvedMime + ';base64,' + base64Data;

  } catch (e) {
    Logger.log('Error getting file as data URI "' + fileName + '": ' + e.toString());
    return null;
  }
}

/**
 * Safely retrieves a value from a row by column index.
 * @param {Array} row The data row.
 * @param {number} index The column index.
 * @param {*} defaultValue Default value if not found.
 * @returns {*} The cell value or default.
 */
function getCellValue(row, index, defaultValue) {
  if (defaultValue === undefined) defaultValue = '';
  if (index < 0 || index >= row.length) return defaultValue;
  var value = row[index];
  return (value !== null && value !== undefined) ? value : defaultValue;
}

/**
 * Builds a map of item names to row indices from sheet data.
 * @param {Array<Array>} data The sheet data including headers.
 * @param {number} itemColIndex The column index for item names.
 * @returns {Object} Map of item name -> row index (0-based, excluding header).
 */
function buildItemMap(data, itemColIndex) {
  var itemMap = {};
  for (var i = 1; i < data.length; i++) {
    var itemName = data[i][itemColIndex];
    if (itemName) {
      itemMap[itemName.toString().trim()] = i;
    }
  }
  return itemMap;
}

/**
 * Gets column indices for a set of headers.
 * @param {Array<string>} sheetHeaders The headers from the sheet.
 * @param {Object} headerMap Mapping of source headers to sheet headers.
 * @returns {Object} Map of header name -> column index.
 */
function getColumnIndices(sheetHeaders, headerMap) {
  var indices = {};
  for (var sourceHeader in headerMap) {
    var targetHeader = headerMap[sourceHeader];
    indices[targetHeader] = sheetHeaders.indexOf(targetHeader);
  }
  return indices;
}

/**
 * Shows a toast notification in the spreadsheet.
 * @param {string} message The message to display.
 * @param {string} [title] The toast title (default 'Wine List').
 * @param {number} [duration] Duration in seconds (default 5).
 */
function showToast(message, title, duration) {
  title = title || 'Wine List';
  duration = duration || 5;
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, duration);
}

/**
 * Formats a date for the R365 API.
 * @param {string} isoDate Date in yyyy-MM-dd format.
 * @returns {string} Date in MM/dd/yyyy format.
 */
function formatDateForR365(isoDate) {
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var date = new Date(isoDate + 'T00:00:00');
  return Utilities.formatDate(date, timeZone, 'MM/dd/yyyy');
}
