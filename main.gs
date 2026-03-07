// main.gs

/**
 * @fileoverview Entry point and orchestration for the Wine List Generator.
 * @OnlyCurrentDoc
 */

// ============================================================================
// Menu
// ============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Wine List')
    .addItem('Generate Wine List', 'generateWineList')
    .addSeparator()
    .addItem('Settings', 'showSettingsDialog')
    .addToUi();

  ui.createMenu('Inventory Tools')
    .addItem('Update Purchases', 'showPurchaseDateDialog')
    .addItem('Upload Variance Data (.csv)', 'showVarianceUploadDialog')
    .addSeparator()
    .addItem('Update Report Feed', 'showFeedUploadDialog')
    .addToUi();
}

// ============================================================================
// Asset Loading
// ============================================================================

/**
 * Fetches all static assets required for the wine list HTML.
 * Collects font files referenced by any heading style or wine entry setting,
 * deduplicates them, and loads each as a base64 data URI.
 *
 * @returns {Object|null} Asset object or null if any required asset is missing.
 */
function loadAssets() {
  var imageSettings = getImageSettings();
  var headingStyles = getAllHeadingStyles();
  var wineEntry = getWineEntryStyle();

  // Collect all unique font filenames
  var fontFiles = new Set();
  for (var t = 1; t <= MAX_HEADING_TYPE; t++) {
    var hs = headingStyles[t];
    if (hs) {
      fontFiles.add(hs.title.font);
      fontFiles.add(hs.subtext.font);
    }
  }
  fontFiles.add(wineEntry.font);

  // Load images
  var footerImageUri = getFileAsDataUri(imageSettings.footer, 'image/png');
  var logoImageUri = getFileAsDataUri(imageSettings.logo, 'image/png');

  // Load fonts
  var fontUris = new Map();
  var missing = [];

  if (!footerImageUri) missing.push('Footer image (' + imageSettings.footer + ')');
  if (!logoImageUri) missing.push('Logo image (' + imageSettings.logo + ')');

  fontFiles.forEach(function(fileName) {
    var ext = extensionOf(fileName);
    var uri = getFileAsDataUri(fileName, getMimeForExtension(ext));
    if (uri) {
      fontUris.set(fileName, {
        uri: uri,
        format: getFontFormatForExtension(ext)
      });
    } else {
      missing.push('Font (' + fileName + ')');
    }
  });

  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert(
      'Error: The following assets could not be found in Google Drive:\n• ' +
      missing.join('\n• ') + '\n\n' +
      'Tip: Use Wine List > Settings to upload fonts or update file assignments.'
    );
    return null;
  }

  return {
    footerImageUri: footerImageUri,
    logoImageUri: logoImageUri,
    fontUris: fontUris
  };
}

// ============================================================================
// Main Generation Flow
// ============================================================================

function generateWineList() {
  var ui = SpreadsheetApp.getUi();

  // Step 1: Load assets
  var assets = loadAssets();
  if (!assets) return;

  // Step 2: Prepare data
  var data = prepareWineListData();
  if (!data) return;

  // Step 3: Warn about orphaned wines
  if (data.orphans.length > 0) {
    var orphanNames = data.orphans.slice(0, 10).map(function(o) {
      return o.name + ' (code: ' + o.code + ')';
    });
    var msg = data.orphans.length + ' wine(s) reference section codes that don\'t exist ' +
      'in the Sections sheet and will be omitted:\n\n• ' + orphanNames.join('\n• ');
    if (data.orphans.length > 10) {
      msg += '\n... and ' + (data.orphans.length - 10) + ' more.';
    }
    var response = ui.alert('Orphaned Wines Found', msg, ui.ButtonSet.OK_CANCEL);
    if (response === ui.Button.CANCEL) return;
  }

  // Step 4: Build brand settings object (now includes all settings groups)
  var pageConfig = getPageConfig();
  var brand = {
    colors:        getColorSettings(),
    welcome:       getWelcomeSettings(),
    headingStyles: getAllHeadingStyles(),
    wineEntry:     getWineEntryStyle(),
    footer:        getFooterSettings(),
    pageConfig:    pageConfig
  };

  // Step 5: Calculate pagination
  showToast('Calculating pagination...');
  var pagination = calculatePagination(data.sections, data.wineMap, brand.headingStyles, pageConfig);

  // Step 6: Build TOC data
  pagination.tocData = buildTOCData(data.sections, pagination.tocPageNumbers);

  // Step 7: Generate HTML
  var wineCount = 0;
  data.wineMap.forEach(function(wines) { wineCount += wines.length; });
  showToast('Generating HTML for ' + wineCount + ' wines across ' +
    pagination.totalPages + ' pages...');

  var html = generateHTML(data.sections, data.wineMap, pagination, assets, brand);

  // Step 8: Show download dialog
  showDownloadDialog(html);
}

// ============================================================================
// Dialog Launchers
// ============================================================================

function showSettingsDialog() {
  var html = HtmlService.createHtmlOutputFromFile('SettingsDialog')
    .setWidth(700)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Wine List Settings');
}

function showPurchaseDateDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var today = new Date();

  var prevMonday = new Date(today);
  var dayOfWeek = today.getDay();
  var daysToSubtract = (dayOfWeek === 0 ? 6 : dayOfWeek - 1) + 7;
  prevMonday.setDate(today.getDate() - daysToSubtract);

  var template = HtmlService.createTemplateFromFile('DateRangeDialog');
  template.defaultStart = Utilities.formatDate(prevMonday, timeZone, 'yyyy-MM-dd');
  template.defaultEnd = Utilities.formatDate(today, timeZone, 'yyyy-MM-dd');

  var html = template.evaluate().setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Purchase Date Range');
}

function showVarianceUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('VarianceUploadDialog')
    .setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Variance Data');
}

function showFeedUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FeedConfigDialog')
    .setWidth(560).setHeight(470);
  SpreadsheetApp.getUi().showModalDialog(html, 'R365 Report Feed Configuration');
}

// Legacy
function updatePurchasesFromR365() { showPurchaseDateDialog(); }
function showUploadDialog() { showFeedUploadDialog(); }
