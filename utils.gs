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
 * Rewrites the width and height attributes of a base64-encoded SVG data URI.
 * Required for Ruxton Bar footer icons because CSS cannot resize images placed
 * via content: url() in @page margin boxes — only the SVG's intrinsic dimensions
 * (its own width/height XML attributes) control the rendered size.
 *
 * Handles both attribute forms:
 *   <svg width="200" height="200" ...>
 *   <svg width="200px" height="200px" ...>
 *
 * If the URI is not an SVG data URI, it is returned unchanged.
 *
 * @param {string} dataUri  A data URI string (data:image/svg+xml;base64,...).
 * @param {number} width    Target width in pixels.
 * @param {number} height   Target height in pixels.
 * @returns {string}        Modified data URI, or original if not SVG / parse fails.
 */
function resizeSvgDataUri_(dataUri, width, height) {
  if (!dataUri || dataUri.indexOf('image/svg+xml') === -1) return dataUri;

  try {
    var base64Part = dataUri.split(',')[1];
    if (!base64Part) return dataUri;

    var svgText = Utilities.newBlob(
      Utilities.base64Decode(base64Part),
      'image/svg+xml'
    ).getDataAsString();

    // Replace existing width / height attributes on the root <svg> element.
    // The regex targets only the opening <svg ...> tag to avoid touching
    // width/height on nested elements (e.g. rect, image).
    svgText = svgText.replace(
      /(<svg\b[^>]*?)\s+width\s*=\s*["'][^"']*["']/i,
      '$1 width="' + width + '"'
    );
    svgText = svgText.replace(
      /(<svg\b[^>]*?)\s+height\s*=\s*["'][^"']*["']/i,
      '$1 height="' + height + '"'
    );

    // If there were no existing width/height attributes, inject them
    // directly after the opening <svg tag.
    if (svgText.indexOf('width="' + width + '"') === -1) {
      svgText = svgText.replace(/<svg\b/, '<svg width="' + width + '" height="' + height + '"');
    }

    var resizedBase64 = Utilities.base64Encode(
      Utilities.newBlob(svgText, 'image/svg+xml').getBytes()
    );
    return 'data:image/svg+xml;base64,' + resizedBase64;

  } catch (e) {
    Logger.log('resizeSvgDataUri_: failed to resize SVG (' + e.toString() + '), using original.');
    return dataUri;
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
