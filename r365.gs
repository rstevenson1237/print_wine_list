// r365.gs

/**
 * @fileoverview Restaurant365 integration for purchases and inventory.
 *
 * R365 connection parameters are stored in DocumentProperties (visible to all
 * editors of the spreadsheet, unlike UserProperties which are per-user).
 *
 * Keys (see R365_DOC_PROPS):
 *   R365_USER          — R365 username / account
 *   R365_FILTER        — Location filter value (GUID or encoded name)
 *   R365_ITEM_CATEGORY — ItemCategory1v2 GUID for wine items
 */

// ============================================================================
// DocumentProperties Keys
// ============================================================================

const R365_DOC_PROPS = {
  USER:          'R365_USER',
  FILTER:        'R365_FILTER',
  ITEM_CATEGORY: 'R365_ITEM_CATEGORY'
};

// ============================================================================
// Config Getters / Setters (called from FeedConfigDialog)
// ============================================================================

/**
 * Returns the current R365 connection parameters from DocumentProperties.
 * Runs a one-time migration from the legacy UserProperties key if needed.
 *
 * @returns {Object} { user, filter, itemCategory, hasConfig }
 */
function getR365Config() {
  migrateR365Config_();

  const p = PropertiesService.getDocumentProperties();
  const user         = p.getProperty(R365_DOC_PROPS.USER)          || '';
  const filter       = p.getProperty(R365_DOC_PROPS.FILTER)        || '';
  const itemCategory = p.getProperty(R365_DOC_PROPS.ITEM_CATEGORY) || '';

  return {
    user,
    filter,
    itemCategory,
    hasConfig: !!(user && filter && itemCategory)
  };
}

/**
 * Saves a single R365 connection parameter to DocumentProperties.
 * Called from FeedConfigDialog when the user edits and saves a field.
 *
 * @param {string} key   One of the R365_DOC_PROPS values.
 * @param {string} value The value to store.
 * @returns {Object} { success: boolean, message: string }
 */
function saveR365Param(key, value) {
  const validKeys = Object.values(R365_DOC_PROPS);
  if (!validKeys.includes(key)) {
    return { success: false, message: 'Unrecognised parameter key: ' + key };
  }
  if (!value || !value.trim()) {
    return { success: false, message: 'Value cannot be empty.' };
  }
  try {
    PropertiesService.getDocumentProperties().setProperty(key, value.trim());
    return { success: true, message: 'Parameter saved.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Returns the fully-built URL template (with date placeholders) for preview.
 * Returns an empty string if the config is incomplete.
 *
 * @returns {string}
 */
function getR365UrlPreview() {
  const config = getR365Config();
  if (!config.hasConfig) return '';
  return buildR365UrlTemplate_(config.user, config.filter, config.itemCategory);
}

// ============================================================================
// URL Builder
// ============================================================================

/**
 * Builds the R365 report URL template from the three stored parameters.
 * The result contains __START_DATE__ and __END_DATE__ placeholders that are
 * substituted at fetch time.
 *
 * @param {string} user         R365 User param value.
 * @param {string} filter       R365 Filter param value.
 * @param {string} itemCategory R365 ItemCategory1v2 param value.
 * @returns {string} Full URL template.
 * @private
 */
function buildR365UrlTemplate_(user, filter, itemCategory) {
  return (
    'https://na02reports.restaurant365.com/ReportServer' +
    '?%2FNA02%2FReceiving%20by%20Purchased%20Item' +
    '&Database=atlasrestaurantgroup' +
    '&User=' + user +
    '&SQLServer=pro-sqlag-571.restaurant365.com' +
    '&TimeZoneCode=EST' +
    '&UtcOffset=-05' +
    '&FilterBy=Location' +
    '&Filter=' + filter +
    '&Start=__START_DATE__' +
    '&End=__END_DATE__' +
    '&DetailLevel=False' +
    '&KeyItems=0' +
    '&Vendor=00000000-0000-0000-0000-000000000000' +
    '&AccountRec=00000000-0000-0000-0000-000000000000' +
    '&Item=00000000-0000-0000-0000-000000000000' +
    '&VendorItemNumber%3Aisnull=True' +
    '&ItemCategory1v2=' + itemCategory +
    '&ItemCategory2v2=00000000-0000-0000-0000-000000000000' +
    '&ItemCategory3v2=00000000-0000-0000-0000-000000000000' +
    '&ParamExpandOrCollapse=Expand' +
    '&UofM=1' +
    '&Transfers=1' +
    '&rs%3AParameterLanguage=' +
    '&rs%3ACommand=Render' +
    '&rs%3AFormat=CSV' +
    '&rc%3AItemPath=Tablix1'
  );
}

// ============================================================================
// Migration
// ============================================================================

/**
 * One-time migration: if the legacy UserProperties R365_URL_TEMPLATE key
 * exists and DocumentProperties are not yet populated, parse the old URL
 * and write the three params to DocumentProperties.
 * @private
 */
function migrateR365Config_() {
  const docProps  = PropertiesService.getDocumentProperties();
  const userProps = PropertiesService.getUserProperties();

  // Already migrated?
  if (docProps.getProperty(R365_DOC_PROPS.USER)) return;

  const oldTemplate = userProps.getProperty('R365_URL_TEMPLATE');
  if (!oldTemplate) return;

  try {
    const params = extractSpecificParams(oldTemplate);
    docProps.setProperties({
      [R365_DOC_PROPS.USER]:          params.user,
      [R365_DOC_PROPS.FILTER]:        params.filter,
      [R365_DOC_PROPS.ITEM_CATEGORY]: params.itemCategory1v2
    });
    userProps.deleteProperty('R365_URL_TEMPLATE');
    Logger.log('R365: migrated config from UserProperties to DocumentProperties.');
  } catch (e) {
    Logger.log('R365 migration skipped: ' + e);
  }
}

// ============================================================================
// Fetch
// ============================================================================

/**
 * Fetches data from the R365 report server.
 * Builds the URL from DocumentProperties parameters.
 *
 * @param {string} startDate Start date in MM/dd/yyyy format.
 * @param {string} endDate   End date in MM/dd/yyyy format.
 * @returns {Array<Array<string>>} Parsed CSV data as 2D array.
 */
function fetchR365Data(startDate, endDate) {
  const config = getR365Config();

  if (!config.hasConfig) {
    throw new Error(
      'R365 connection is not configured. Use "Inventory Tools > Update Report Feed" ' +
      'to enter your Username, Location Filter, and Item Category.'
    );
  }

  const urlTemplate = buildR365UrlTemplate_(config.user, config.filter, config.itemCategory);

  const startDateTime = startDate + ' 00:00:00';
  const endDateTime   = endDate   + ' 00:00:00';

  const fullUrl = urlTemplate
    .replace('__START_DATE__', encodeURIComponent(startDateTime))
    .replace('__END_DATE__',   encodeURIComponent(endDateTime));

  Logger.log('Fetching R365 URL: ' + fullUrl);

  const response     = UrlFetchApp.fetch(fullUrl, { muteHttpExceptions: true });
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    const errorBody = response.getContentText();
    Logger.log('R365 error (' + responseCode + '): ' + errorBody);
    throw new Error('Failed to fetch data. Server responded with code: ' + responseCode);
  }

  const csvContent = response.getContentText();
  Logger.log('Raw CSV received:\n' + csvContent.substring(0, 500));

  // Remove R365's metadata header lines (first 3 lines)
  const lines = csvContent.split('\n');
  if (lines.length < 4) {
    Logger.log('CSV has fewer than 4 lines, returning empty array.');
    return [];
  }

  const cleanedCsv = lines.slice(3).join('\n');
  return Utilities.parseCsv(cleanedCsv);
}

// ============================================================================
// Purchase Update
// ============================================================================

/**
 * Processes R365 purchase data and updates the Data sheet.
 * Called from DateRangeDialog.html.
 *
 * @param {string} startDateISO Start date in yyyy-MM-dd format.
 * @param {string} endDateISO   End date in yyyy-MM-dd format.
 * @returns {Object} Result object with success status and message.
 */
function processDatesAndRunUpdate(startDateISO, endDateISO) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    showToast('Starting purchase update from R365...', 'Inventory Tools');

    const startDate = formatDateForR365(startDateISO);
    const endDate   = formatDateForR365(endDateISO);

    const csvData = fetchR365Data(startDate, endDate);
    if (!csvData || csvData.length < 2) {
      return {
        success: true,
        message: 'No purchase data found for: ' + startDate + ' – ' + endDate
      };
    }

    const dataSheet = ss.getSheetByName(SHEETS.DATA);
    if (!dataSheet) throw new Error('Sheet "Data" not found.');

    const listSheet = ss.getSheetByName(SHEETS.LIST);
    const result    = updateDataFromCSV(dataSheet, listSheet, csvData, R365_HEADER_MAP);

    let message = 'Update complete.\n\nItems Updated: ' + result.updated;
    if (result.added.length > 0) {
      message += '\n\nItems Added: ' + result.added.length + '\n• ' + result.added.join('\n• ');
    } else {
      message += '\n\nNo new items were added.';
    }

    return { success: true, message };

  } catch (e) {
    Logger.log('Purchase update error: ' + e);
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================================
// Variance CSV
// ============================================================================

/**
 * Processes variance CSV data and updates the Data sheet.
 * Called from VarianceUploadDialog.html.
 *
 * @param {string} csvContent Raw CSV content as string.
 * @returns {Object} Result object with success status and message.
 */
function processVarianceCSV(csvContent) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SHEETS.DATA);

  if (!dataSheet) {
    return { success: false, message: 'Error: Sheet "Data" not found.' };
  }

  try {
    const csvData = Utilities.parseCsv(csvContent);
    if (csvData.length < 2) {
      return { success: true, message: 'File is empty or contains only headers.' };
    }

    const csvHeaders     = csvData[0];
    const missingColumns = Object.keys(VARIANCE_HEADER_MAP).filter(h => !csvHeaders.includes(h));
    if (missingColumns.length > 0) {
      return {
        success: false,
        message: 'Uploaded file is missing required columns: ' + missingColumns.join(', ')
      };
    }

    const listSheet = ss.getSheetByName(SHEETS.LIST);
    clearCountColumn(dataSheet);

    const result = updateDataFromCSV(
      dataSheet, listSheet, csvData, VARIANCE_HEADER_MAP, { filterCategory: 'Wine' }
    );

    let message = 'Upload Complete.\n\nItems Updated: ' + result.updated;
    message += result.added.length > 0
      ? '\nItems Added: ' + result.added.length + '\n• ' + result.added.join('\n• ')
      : '\nItems Added: 0';

    return { success: true, message };

  } catch (e) {
    Logger.log('Variance processing error: ' + e);
    return { success: false, message: 'Error: ' + e.message };
  }
}

/**
 * Clears the Count column in the Data sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function clearCountColumn(sheet) {
  const headers    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const countIndex = headers.indexOf('Count');
  if (countIndex === -1) throw new Error('Column "Count" not found in Data sheet.');

  const numRows = sheet.getLastRow() - 1;
  if (numRows > 0) {
    sheet.getRange(2, countIndex + 1, numRows, 1).clearContent();
  }
}

// ============================================================================
// Toast Pricing CSV
// ============================================================================

/**
 * Processes a Toast pricing CSV export and updates Data column B (Name/Price).
 * Called from ToastPricingUploadDialog.html.
 *
 * Rows are included only when:
 *   - Archived = "No"
 *   - Modifier = "No"
 *   - Name starts with "BTL "
 *
 * The "BTL " prefix is stripped before matching against Data!A2:A.
 * Data!B2:B is cleared first; duplicate matches within the file are
 * flagged and returned to the user for investigation.
 *
 * @param {string} csvContent Raw CSV content as string.
 * @returns {Object} { success, message, duplicates: string[] }
 */
function processToastPricingCSV(csvContent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SHEETS.DATA);

  if (!dataSheet) {
    return { success: false, message: 'Error: Sheet "Data" not found.', duplicates: [] };
  }

  try {
    const csvData = Utilities.parseCsv(csvContent);
    if (csvData.length < 2) {
      return { success: true, message: 'File is empty or contains only headers.', duplicates: [] };
    }

    // ── Validate required columns ─────────────────────────────────────────
    const csvHeaders     = csvData[0];
    const requiredCols   = Object.keys(TOAST_PRICING_HEADER_MAP);
    const missingColumns = requiredCols.filter(h => !csvHeaders.includes(h));
    if (missingColumns.length > 0) {
      return {
        success:    false,
        message:    'Uploaded file is missing required columns: ' + missingColumns.join(', '),
        duplicates: []
      };
    }

    const nameIdx      = csvHeaders.indexOf('Name');
    const priceIdx     = csvHeaders.indexOf('Base Price');
    const archivedIdx  = csvHeaders.indexOf('Archived');
    const modifierIdx  = csvHeaders.indexOf('Modifier');

    // ── Load Data!A and reset Data!B ─────────────────────────────────────
    const lastRow = dataSheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, message: 'Data sheet has no data rows.', duplicates: [] };
    }

    const numRows      = lastRow - 1;
    const itemColValues = dataSheet.getRange(2, 1, numRows, 1).getValues(); // A2:A

    // Clear column B now
    dataSheet.getRange(2, 2, numRows, 1).clearContent();

    // Build a lookup map: lowercase item name → 0-based row index into itemColValues
    const itemIndex = {};
    itemColValues.forEach(function(row, i) {
      const key = (row[0] || '').toString().trim().toLowerCase();
      if (key) itemIndex[key] = i;
    });

    // Working array for column B output (parallel to itemColValues)
    const outputB = new Array(numRows).fill('');

    // ── Process CSV rows ──────────────────────────────────────────────────
    let updated    = 0;
    const duplicates = [];
    const notFound   = [];

    for (let r = 1; r < csvData.length; r++) {
      const row = csvData[r];
      if (!row || row.length < csvHeaders.length) continue;

      const archived = (row[archivedIdx]  || '').toString().trim();
      const modifier = (row[modifierIdx]  || '').toString().trim();
      const rawName  = (row[nameIdx]       || '').toString().trim();

      // Apply filters
      if (archived.toLowerCase() !== 'no')   continue;
      if (modifier.toLowerCase() !== 'no')   continue;
      if (!rawName.startsWith('BTL '))        continue;

      const strippedName = rawName.substring(4).trim(); // remove "BTL "
      const lookupKey    = strippedName.toLowerCase();
      const rawPrice     = (row[priceIdx]  || '').toString().trim();

      if (!(lookupKey in itemIndex)) {
        notFound.push(strippedName);
        continue;
      }

      const idx = itemIndex[lookupKey];

      if (outputB[idx] !== '') {
        // Slot already written in this run — record both the previous value
        // and the current one so the user can see what's conflicting
        duplicates.push(
          '"' + strippedName + '" — previous: ' + outputB[idx] + ', new: ' + rawPrice
        );
        outputB[idx] = rawPrice; // overwrite with latest; user can investigate
      } else {
        outputB[idx] = rawPrice;
        updated++;
      }
    }

    // ── Write column B back in one batch ─────────────────────────────────
    const writeRange = dataSheet.getRange(2, 2, numRows, 1);
    writeRange.setValues(outputB.map(function(v) { return [v]; }));

    // ── Build result message ──────────────────────────────────────────────
    let message = 'Upload Complete.\n\nItems Updated: ' + updated;

    if (notFound.length > 0) {
      message += '\nNot Matched in Data: ' + notFound.length;
    }

    if (duplicates.length > 0) {
      message += '\n\n⚠️ Duplicates found: ' + duplicates.length +
                 '\nThese items appeared more than once in the file. ' +
                 'See the list below for details.';
    }

    return { success: true, message, duplicates, notFound };

  } catch (e) {
    Logger.log('Toast pricing processing error: ' + e);
    return { success: false, message: 'Error: ' + e.message, duplicates: [] };
  }
}

// ============================================================================
// Feed Import (.atomsvc)
// ============================================================================

/**
 * Processes an uploaded .atomsvc file, extracts the three connection params,
 * and saves them to DocumentProperties.
 * Returns the extracted config so the dialog can populate fields immediately.
 *
 * @param {string} xmlContent The .atomsvc file content.
 * @returns {Object} { success, message, config?: { user, filter, itemCategory } }
 */
function processUploadedFeed(xmlContent) {
  try {
    const uploadedDoc  = XmlService.parse(xmlContent);
    const uploadedRoot = uploadedDoc.getRootElement();
    const appNs        = XmlService.getNamespace('http://www.w3.org/2007/app');

    const uploadedCollection = uploadedRoot
      .getChild('workspace', appNs)
      .getChild('collection', appNs);

    if (!uploadedCollection) {
      throw new Error('Invalid file structure: <collection> tag not found.');
    }

    const uploadedHref = uploadedCollection.getAttribute('href').getValue();
    if (!uploadedHref) {
      throw new Error('Invalid file structure: "href" attribute not found.');
    }

    const params = extractSpecificParams(uploadedHref);

    // Save to DocumentProperties
    PropertiesService.getDocumentProperties().setProperties({
      [R365_DOC_PROPS.USER]:          params.user,
      [R365_DOC_PROPS.FILTER]:        params.filter,
      [R365_DOC_PROPS.ITEM_CATEGORY]: params.itemCategory1v2
    });

    Logger.log('R365 params saved from .atomsvc — User: ' + params.user +
               ', Filter: ' + params.filter +
               ', ItemCategory: ' + params.itemCategory1v2);

    return {
      success: true,
      message: 'Feed imported. Username, Location Filter, and Item Category have been updated.',
      config: {
        user:         params.user,
        filter:       params.filter,
        itemCategory: params.itemCategory1v2
      }
    };

  } catch (e) {
    Logger.log('Feed processing error: ' + e);
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================================
// URL Param Extraction
// ============================================================================

/**
 * Extracts User, Filter, and ItemCategory1v2 parameters from a URL.
 *
 * @param {string} url
 * @returns {{ user: string, filter: string, itemCategory1v2: string }}
 */
function extractSpecificParams(url) {
  const params = { user: '', filter: '', itemCategory1v2: '' };

  const userMatch    = url.match(/[?&]User=([^&]+)/);
  const filterMatch  = url.match(/[?&]Filter=([^&]+)/);
  const itemCatMatch = url.match(/[?&]ItemCategory1v2=([^&]+)/);

  if (userMatch)    params.user            = userMatch[1];
  if (filterMatch)  params.filter          = filterMatch[1];
  if (itemCatMatch) params.itemCategory1v2 = itemCatMatch[1];

  if (!params.user || !params.filter || !params.itemCategory1v2) {
    throw new Error(
      'Could not extract required parameters (User, Filter, ItemCategory1v2) from the uploaded file.'
    );
  }

  return params;
}

// ============================================================================
// Data Sheet Update (shared by R365 and Variance flows)
// ============================================================================

/**
 * Updates the Data sheet from CSV data using the provided header mapping.
 * Also adds new items to the List sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dataSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} listSheet (optional)
 * @param {Array<Array>} csvData Parsed CSV data.
 * @param {Object} headerMap Mapping of CSV headers to sheet headers.
 * @param {Object} options e.g. { filterCategory: 'Wine' }
 * @returns {{ updated: number, added: string[] }}
 */
function updateDataFromCSV(dataSheet, listSheet, csvData, headerMap, options = {}) {
  const csvHeaders  = csvData[0];
  const sheetData   = dataSheet.getDataRange().getValues();
  const sheetHeaders = sheetData[0];

  const itemColIndex = sheetHeaders.indexOf('Item');
  if (itemColIndex === -1) throw new Error('Column "Item" not found in Data sheet.');

  const itemMap = buildItemMap(sheetData, itemColIndex);

  let listHeaders    = [];
  let listItemIndex  = -1;
  let listNotesIndex = -1;

  if (listSheet) {
    listHeaders    = listSheet.getRange(1, 1, 1, listSheet.getLastColumn()).getValues()[0];
    listItemIndex  = listHeaders.indexOf('Item');
    listNotesIndex = listHeaders.indexOf('Notes');
  }

  const csvItemIndex = csvHeaders.indexOf(
    Object.keys(headerMap).find(k => headerMap[k] === 'Item')
  );
  const csvCat1Index = options.filterCategory
    ? csvHeaders.indexOf(Object.keys(headerMap).find(k => headerMap[k] === 'Category1'))
    : -1;

  if (options.filterCategory && csvCat1Index === -1) {
    throw new Error("Filter requested on 'Category1' but no matching column found in CSV headers.");
  }

  const dataRowsToAppend = [];
  const listRowsToAppend = [];
  const newItems         = [];
  const batchItemMap     = {};
  let updatedCount       = 0;

  for (let i = 1; i < csvData.length; i++) {
    const csvRow = csvData[i];

    if (options.filterCategory && csvCat1Index > -1) {
      const cat1 = (csvRow[csvCat1Index] || '').toString().trim();
      if (cat1 !== options.filterCategory) continue;
    }

    const itemName = csvRow[csvItemIndex];
    if (!itemName) continue;

    const trimmedName      = itemName.toString().trim();
    const existingRowIndex = itemMap[trimmedName];

    if (existingRowIndex !== undefined) {
      for (const csvHeader in headerMap) {
        const sheetHeader  = headerMap[csvHeader];
        const sheetColIndex = sheetHeaders.indexOf(sheetHeader);
        const csvColIndex   = csvHeaders.indexOf(csvHeader);
        if (sheetColIndex > -1 && csvColIndex > -1 && csvRow[csvColIndex] !== undefined) {
          dataSheet.getRange(existingRowIndex + 1, sheetColIndex + 1).setValue(csvRow[csvColIndex]);
        }
      }
      updatedCount++;

    } else if (batchItemMap[trimmedName] !== undefined) {
      const batchIndex = batchItemMap[trimmedName];
      for (const csvHeader in headerMap) {
        const sheetHeader   = headerMap[csvHeader];
        const sheetColIndex = sheetHeaders.indexOf(sheetHeader);
        const csvColIndex   = csvHeaders.indexOf(csvHeader);
        if (sheetColIndex > -1 && csvColIndex > -1 && csvRow[csvColIndex] !== undefined) {
          dataRowsToAppend[batchIndex][sheetColIndex] = csvRow[csvColIndex];
        }
      }

    } else {
      const newRow = new Array(sheetHeaders.length).fill('');
      for (const csvHeader in headerMap) {
        const sheetHeader   = headerMap[csvHeader];
        const sheetColIndex = sheetHeaders.indexOf(sheetHeader);
        const csvColIndex   = csvHeaders.indexOf(csvHeader);
        if (sheetColIndex > -1 && csvColIndex > -1 && csvRow[csvColIndex] !== undefined) {
          newRow[sheetColIndex] = csvRow[csvColIndex];
        }
      }
      batchItemMap[trimmedName] = dataRowsToAppend.length;
      dataRowsToAppend.push(newRow);
      newItems.push(trimmedName);

      if (listSheet && listItemIndex > -1 && listNotesIndex > -1) {
        const listRow = new Array(listHeaders.length).fill('');
        listRow[listItemIndex]  = trimmedName;
        listRow[listNotesIndex] = 'NEW';
        listRowsToAppend.push(listRow);
      }
    }
  }

  if (dataRowsToAppend.length > 0) {
    dataSheet.getRange(
      dataSheet.getLastRow() + 1, 1, dataRowsToAppend.length, sheetHeaders.length
    ).setValues(dataRowsToAppend);
  }

  if (listSheet && listRowsToAppend.length > 0) {
    listSheet.getRange(
      listSheet.getLastRow() + 1, 1, listRowsToAppend.length, listHeaders.length
    ).setValues(listRowsToAppend);
  }

  return { updated: updatedCount, added: newItems };
}
