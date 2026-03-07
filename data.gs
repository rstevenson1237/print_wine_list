// data.gs

/**
 * @fileoverview Data fetching and preparation for the Wine List Generator.
 *
 * The data pipeline:
 *   1. Read the Sections sheet (flat outline: Code, Type, Title, Subtext)
 *   2. Read the List sheet and build a Map<code, Array<wine>>
 *   3. Validate that every wine references a valid section code
 *   4. Return both structures for the pagination and rendering loops
 */

// ============================================================================
// Public API
// ============================================================================

/**
 * Prepares the complete wine list data for HTML generation.
 * This is the main data pipeline function called from main.gs.
 *
 * @returns {Object|null} { sections, wineMap, orphans } or null on error.
 *   sections  — ordered Array of section row objects from the Sections sheet
 *   wineMap   — Map<number, Array<Object>> keyed by section code
 *   orphans   — Array of { name, code } for wines whose code has no section
 */
function prepareWineListData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName(SHEETS.LIST);
  const sectionsSheet = ss.getSheetByName(SHEETS.SECTIONS);

  if (!listSheet || !sectionsSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets ("List" and "Sections") not found.');
    return null;
  }

  // Step 1: Read sections
  const sections = getSectionData(sectionsSheet);
  if (sections.length === 0) {
    SpreadsheetApp.getUi().alert('No section data found in the Sections sheet.');
    return null;
  }

  // Step 2: Read wines and group by code
  const wineData = getWineData(listSheet);
  if (wineData.length === 0) {
    SpreadsheetApp.getUi().alert('No wine data found or missing required columns in the List sheet.');
    return null;
  }
  const wineMap = buildWineMap(wineData);

  // Step 3: Validate — find orphaned wines
  const sectionCodes = new Set(sections.filter(function(s) { return s.code > 0; }).map(function(s) { return s.code; }));
  const orphans = [];
  wineMap.forEach(function(wines, code) {
    if (!sectionCodes.has(code)) {
      wines.forEach(function(w) {
        orphans.push({ name: w.name, code: code });
      });
    }
  });

  Logger.log('Data pipeline: ' + sections.length + ' sections, ' +
    wineData.length + ' wines, ' + orphans.length + ' orphans');

  return {
    sections: sections,
    wineMap: wineMap,
    orphans: orphans
  };
}

// ============================================================================
// Sections Sheet Reader
// ============================================================================

/**
 * Reads the Sections sheet and returns an ordered array of section objects.
 *
 * Columns:
 *   A — Section code (numeric, links to Data sheet)
 *   B — Heading type (1–4, matches SECTION_TYPES)
 *   C — Title text
 *   D — Subtext (optional)
 *   E — Force new page (checkbox / any truthy value)
 *
 * @returns {Array<Object>} Ordered section descriptors:
 *   { code, type, title, subtext, forceNewPage }
 */
function getSectionData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SECTIONS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "' + SHEETS.SECTIONS + '" not found.');
    return null;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];   // header row only

  var sections = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var code        = parseInt(row[0]) || 0;
    var typeRaw     = row[1];
    var title       = (row[2] || '').toString().trim();
    var subtext     = (row[3] || '').toString().trim();
    var forceRaw    = row[4];

    if (!title) continue;   // skip blank rows

    // Resolve type: accept numeric (1–4) or label string ('STYLE', 'COUNTRY', etc.)
    var typeLevel = resolveType_(typeRaw);
    if (!typeLevel) {
      Logger.log('Sections row ' + (i + 1) + ': unrecognised type "' + typeRaw + '" — skipped.');
      continue;
    }

    // Column E: checkbox (TRUE/FALSE), or any truthy string ('yes', 'x', '1', etc.)
    var forceNewPage = false;
    if (typeof forceRaw === 'boolean') {
      forceNewPage = forceRaw;
    } else if (forceRaw !== null && forceRaw !== undefined && forceRaw !== '') {
      var str = forceRaw.toString().trim().toLowerCase();
      forceNewPage = (str === 'true' || str === '1' || str === 'yes' || str === 'x');
    }

    sections.push({
      code:        code,
      type:        typeLevel,
      title:       title,
      subtext:     subtext,
      forceNewPage: forceNewPage
    });
  }

  return sections;
}

/**
 * Resolves a raw cell value from the Type column to a numeric level (1–4).
 * Accepts integers, numeric strings, or the label strings from SECTION_TYPES.
 *
 * @param {*} raw Raw cell value.
 * @returns {number|null} Type level, or null if unrecognised.
 * @private
 */
function resolveType_(raw) {
  if (!raw && raw !== 0) return null;

  // Numeric or numeric string
  var asInt = parseInt(raw);
  if (!isNaN(asInt) && asInt >= 1 && asInt <= MAX_HEADING_TYPE) return asInt;

  // Label string ('STYLE', 'COUNTRY', etc.) — reverse lookup in SECTION_TYPES
  var upper = raw.toString().trim().toUpperCase();
  for (var level in SECTION_TYPES) {
    if (SECTION_TYPES[level].toUpperCase() === upper) return parseInt(level);
  }

  return null;
}

// ============================================================================
// List Sheet Reader
// ============================================================================

/**
 * Reads the List sheet and returns an array of wine objects.
 * Uses header-based column detection for resilience.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The List sheet.
 * @returns {Array<Object>} Array of wine objects.
 */
function getWineData(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastCol ? sheet.getLastColumn() : sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 3) return [];

  var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = allData[0].map(function(h) { return h.toString().trim(); });

  // Find required columns
  var idx = {
    name:            headers.indexOf('Name'),
    regionCode:      headers.indexOf('Section'),
    vintage:         headers.indexOf('Vintage'),
    price:           headers.indexOf('MenuPrice'),
    bin:             headers.indexOf('Bin'),
    cost:            headers.indexOf('Cost'),
    uofm:            headers.indexOf('UofM'),
    storageLocation: headers.indexOf('StorageLocation'),
    newPrice:        headers.indexOf('NewPrice'),
    toastPrice:      headers.indexOf('ToastPrice'),
    count:           headers.indexOf('Count'),
    vendor:          headers.indexOf('Vendor'),
    par:             headers.indexOf('Par'),
    lastPurchase:    headers.indexOf('LastPurchase'),
    order:           headers.indexOf('Order'),
    notes:           headers.indexOf('Notes')
  };

  if (idx.name < 0 || idx.regionCode < 0) {
    Logger.log('Missing required columns: Name=' + idx.name + ', Section=' + idx.regionCode);
    return [];
  }

  var wines = [];

  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var wineName = getCellValue(row, idx.name, '').toString().trim();
    var notes = getCellValue(row, idx.notes, '').toString().trim();

    // Skip entries with notes (hidden/86'd items) or no name
    if (notes || !wineName) continue;

    // Parse the section code: handle both "107-Beaujolais" and plain "107"
    var rawSection = getCellValue(row, idx.regionCode, '').toString().trim();
    var sectionCode = parseInt(rawSection) || 0;

    wines.push({
      name: wineName,
      sectionCode: sectionCode,
      vintage: getCellValue(row, idx.vintage).toString().trim(),
      price: parseFloat(getCellValue(row, idx.price)) || 0,
      bin: getCellValue(row, idx.bin).toString().trim(),
      cost: parseFloat(getCellValue(row, idx.cost)) || 0,
      uofm: getCellValue(row, idx.uofm).toString().trim(),
      storageLocation: getCellValue(row, idx.storageLocation).toString().trim(),
      newPrice: parseFloat(getCellValue(row, idx.newPrice)) || 0,
      toastPrice: parseFloat(getCellValue(row, idx.toastPrice)) || 0,
      count: parseInt(getCellValue(row, idx.count)) || 0,
      vendor: getCellValue(row, idx.vendor).toString().trim(),
      par: parseInt(getCellValue(row, idx.par)) || 0,
      lastPurchase: getCellValue(row, idx.lastPurchase),
      order: getCellValue(row, idx.order),
      notes: notes
    });
  }

  return wines;
}

// ============================================================================
// Wine Map Builder
// ============================================================================

/**
 * Groups wines by their section code into a Map.
 * Wines within each group are sorted alphabetically by name.
 *
 * @param {Array<Object>} wineData Array of wine objects from getWineData.
 * @returns {Map<number, Array<Object>>} Map of sectionCode → sorted wine array.
 */
function buildWineMap(wineData) {
  var map = new Map();

  wineData.forEach(function(wine) {
    var code = wine.sectionCode;
    if (!map.has(code)) {
      map.set(code, []);
    }
    map.get(code).push(wine);
  });

  // Sort each group alphabetically by name
  map.forEach(function(wines) {
    wines.sort(function(a, b) { return a.name.localeCompare(b.name); });
  });

  return map;
}
