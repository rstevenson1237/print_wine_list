// settings.gs

/**
 * @fileoverview Settings management for configurable Wine List options.
 *
 * All settings are stored in DocumentProperties so that every editor of
 * a given spreadsheet copy sees the same configuration. Each property has
 * its own copy of the sheet; all managers at that property share one config.
 *
 * Setting groups:
 *   Brand          — active brand preset name
 *   Page           — page size, margins, buffer
 *   Heading Styles — per-Type title and subtext styling
 *   Wine Entry     — font, size, color, weight, style
 *   Images         — logo and footer image file assignments
 *   Colors         — primary brand color and text color
 *   Welcome        — title page toggle, date, three text lines
 *   Footer         — footer style, running label options
 *
 * Asset files (fonts + images) live in the spreadsheet's Drive folder.
 * Filenames are stored in DocumentProperties as assignments; the binary
 * content is never stored in properties — it is read from Drive at use time.
 */

// ============================================================================
// PropertiesService Key Helpers
// ============================================================================

function headingKey_(type, group, prop) {
  return 'TYPE_' + type + '_' + group + '_' + prop;
}

const PROP_KEYS = {
  BRAND_NAME: 'BRAND_NAME',

  // Page
  PAGE_SIZE_PRESET:   'PAGE_SIZE_PRESET',
  PAGE_WIDTH:         'PAGE_WIDTH',
  PAGE_HEIGHT:        'PAGE_HEIGHT',
  PAGE_MARGIN_TOP:    'PAGE_MARGIN_TOP',
  PAGE_MARGIN_BOTTOM: 'PAGE_MARGIN_BOTTOM',
  PAGE_MARGIN_INNER:  'PAGE_MARGIN_INNER',
  PAGE_MARGIN_OUTER:  'PAGE_MARGIN_OUTER',
  PAGE_BUFFER:        'PAGE_BUFFER',
  FRONT_MATTER_PAGES: 'FRONT_MATTER_PAGES',

  // Images (stores filenames only — binary content stays in Drive)
  IMAGE_LOGO:   'IMAGE_LOGO_FILE',
  IMAGE_FOOTER: 'IMAGE_FOOTER_FILE',

  // Colors
  COLOR_PRIMARY: 'COLOR_PRIMARY',
  COLOR_TEXT:    'COLOR_TEXT',

  // Welcome / Title Page
  SHOW_TITLE_PAGE: 'SHOW_TITLE_PAGE',
  SHOW_DATE:       'SHOW_DATE',
  WELCOME_LINE1:   'WELCOME_LINE1',
  WELCOME_LINE2:   'WELCOME_LINE2',
  WELCOME_LINE3:   'WELCOME_LINE3',

  // Wine Entry
  WINE_FONT:   'WINE_ENTRY_FONT',
  WINE_SIZE:   'WINE_ENTRY_SIZE',
  WINE_COLOR:  'WINE_ENTRY_COLOR',
  WINE_WEIGHT: 'WINE_ENTRY_WEIGHT',
  WINE_STYLE:  'WINE_ENTRY_STYLE',

  // Footer
  FOOTER_PAGE_NUMBER_POSITION: 'FOOTER_PAGE_NUMBER_POSITION',
  FOOTER_RULE:                 'FOOTER_RULE',
  SHOW_RUNNING_LABEL:          'SHOW_RUNNING_LABEL',
  RUNNING_LABEL_POSITION:      'RUNNING_LABEL_POSITION'
};

// ============================================================================
// Brand Preset Management
// ============================================================================

/**
 * Returns the active brand name from DocumentProperties.
 * @returns {string}
 */
function getBrandName(allProps) {
  if (allProps) return allProps[PROP_KEYS.BRAND_NAME] || 'BYGONE';
  return PropertiesService.getDocumentProperties().getProperty(PROP_KEYS.BRAND_NAME) || 'BYGONE';
}

/**
 * Returns the active brand preset object from config.gs.
 * @returns {Object}
 */
function getActiveBrandPreset(allProps) {
  var name = getBrandName(allProps);
  return BRAND_PRESETS[name] || BRAND_PRESETS.BYGONE;
}

/**
 * Loads a brand preset, writing all its values into DocumentProperties.
 * Overwrites all current settings with preset defaults.
 *
 * Note: deleteAllProperties() is intentionally NOT called here. DocumentProperties
 * also holds R365 connection parameters (and any future document-scoped keys) that
 * must not be disturbed. Every settings key is explicitly overwritten below, so no
 * stale settings keys can survive a preset load.
 *
 * @param {string} brandName  Key from BRAND_PRESETS (e.g., 'BYGONE', 'RUXTON').
 * @returns {Object} { success, message }
 */
function loadBrandPreset(brandName) {
  var preset = BRAND_PRESETS[brandName];
  if (!preset) {
    return { success: false, message: 'Unknown brand preset: ' + brandName };
  }

  try {
    var p = PropertiesService.getDocumentProperties();
    var props = {};

    props[PROP_KEYS.BRAND_NAME] = brandName;

    // Page
    props[PROP_KEYS.PAGE_SIZE_PRESET]   = preset.page.sizePreset;
    props[PROP_KEYS.PAGE_WIDTH]         = preset.page.width.toString();
    props[PROP_KEYS.PAGE_HEIGHT]        = preset.page.height.toString();
    props[PROP_KEYS.PAGE_MARGIN_TOP]    = preset.page.marginTop.toString();
    props[PROP_KEYS.PAGE_MARGIN_BOTTOM] = preset.page.marginBottom.toString();
    props[PROP_KEYS.PAGE_MARGIN_INNER]  = preset.page.marginInner.toString();
    props[PROP_KEYS.PAGE_MARGIN_OUTER]  = preset.page.marginOuter.toString();
    props[PROP_KEYS.PAGE_BUFFER]        = preset.page.pageBuffer.toString();
    props[PROP_KEYS.FRONT_MATTER_PAGES] = preset.frontMatterPages.toString();

    // Colors
    props[PROP_KEYS.COLOR_PRIMARY] = preset.colors.primary;
    props[PROP_KEYS.COLOR_TEXT]    = preset.colors.text;

    // Images (filename assignments only — files must already exist in Drive)
    props[PROP_KEYS.IMAGE_LOGO]   = preset.images.logo;
    props[PROP_KEYS.IMAGE_FOOTER] = preset.images.footer;

    // Welcome / Title Page
    props[PROP_KEYS.SHOW_TITLE_PAGE] = preset.welcome.showTitlePage.toString();
    props[PROP_KEYS.SHOW_DATE]       = preset.welcome.showDate.toString();
    props[PROP_KEYS.WELCOME_LINE1]   = preset.welcome.line1;
    props[PROP_KEYS.WELCOME_LINE2]   = preset.welcome.line2;
    props[PROP_KEYS.WELCOME_LINE3]   = preset.welcome.line3;

    // Wine Entry
    props[PROP_KEYS.WINE_FONT]   = preset.wineEntry.font;
    props[PROP_KEYS.WINE_SIZE]   = preset.wineEntry.size.toString();
    props[PROP_KEYS.WINE_COLOR]  = preset.wineEntry.color;
    props[PROP_KEYS.WINE_WEIGHT] = preset.wineEntry.weight;
    props[PROP_KEYS.WINE_STYLE]  = preset.wineEntry.style;

    // Footer
    props[PROP_KEYS.FOOTER_PAGE_NUMBER_POSITION] = preset.footer.pageNumberPosition;
    props[PROP_KEYS.FOOTER_RULE]                 = preset.footer.footerRule;
    props[PROP_KEYS.SHOW_RUNNING_LABEL]          = preset.footer.showRunningLabel.toString();
    props[PROP_KEYS.RUNNING_LABEL_POSITION]      = preset.footer.runningLabelPosition;

    // Heading Styles
    for (var t = 1; t <= MAX_HEADING_TYPE; t++) {
      var hs = preset.headingStyles[t];
      if (hs && hs.title) {
        Object.keys(hs.title).forEach(function(k) {
          props[headingKey_(t, 'TITLE', k.toUpperCase())] = hs.title[k].toString();
        });
      }
      if (hs && hs.subtext) {
        Object.keys(hs.subtext).forEach(function(k) {
          props[headingKey_(t, 'SUB', k.toUpperCase())] = hs.subtext[k].toString();
        });
      }
    }

    p.setProperties(props);
    return { success: true, message: 'Brand preset "' + preset.label + '" loaded. All settings updated.' };
  } catch (e) {
    return { success: false, message: 'Error loading preset: ' + e.toString() };
  }
}

/**
 * Saves the brand name only (without loading full preset).
 * @param {string} brandName
 * @returns {Object}
 */
function saveBrandName(brandName) {
  PropertiesService.getDocumentProperties().setProperty(PROP_KEYS.BRAND_NAME, brandName);
  return { success: true, message: 'Brand set to ' + brandName };
}

// ============================================================================
// Asset Discovery — Fonts & Images
// ============================================================================

/**
 * Scans the spreadsheet's Drive folder once, returning both font and image
 * filename lists in a single pass. Eliminates the four separate folder
 * iterations that previously occurred during getAllSettings() and
 * getAssetPreviewData().
 * @private
 * @returns {{ fonts: string[], images: string[] }}
 */
function scanDriveFolder_() {
  var fonts  = [];
  var images = [];
  try {
    var ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var parents = ssFile.getParents();
    if (!parents.hasNext()) return { fonts: fonts, images: images };
    var folder    = parents.next();
    var fontExts  = ['otf', 'ttf', 'woff', 'woff2'];
    var imageExts = ['png', 'jpg', 'jpeg', 'svg', 'webp'];
    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      var ext  = name.split('.').pop().toLowerCase();
      if (fontExts.indexOf(ext)  > -1) fonts.push(name);
      if (imageExts.indexOf(ext) > -1) images.push(name);
    }
    fonts.sort();
    images.sort();
  } catch (e) {
    Logger.log('scanDriveFolder_: ' + e.toString());
  }
  return { fonts: fonts, images: images };
}

/**
 * Scans the spreadsheet's Drive folder for font files.
 * Returns a sorted array of filenames.
 * @returns {string[]}
 */
function getAvailableFonts() {
  return scanDriveFolder_().fonts;
}

/**
 * Scans the spreadsheet's Drive folder for image files.
 * Returns a sorted array of filenames.
 * @returns {string[]}
 */
function getAvailableImages() {
  return scanDriveFolder_().images;
}

/**
 * Loads every font and image in the Drive folder as a base64 data URI.
 * Called once on Settings dialog open to enable live previews without
 * further server round trips. Fonts and images are treated identically here.
 *
 * @returns {{ fonts: Object.<string,{uri:string,format:string}>, images: Object.<string,string> }}
 */
function getAssetPreviewData() {
  var assets = scanDriveFolder_();

  var imageMimeMap = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
    svg: 'image/svg+xml', webp: 'image/webp'
  };

  var fonts  = {};
  var images = {};

  assets.fonts.forEach(function(fileName) {
    var ext = extensionOf(fileName);
    var uri = getFileAsDataUri(fileName, getMimeForExtension(ext));
    if (uri) fonts[fileName] = { uri: uri, format: getFontFormatForExtension(ext) };
  });

  assets.images.forEach(function(fileName) {
    var ext  = extensionOf(fileName);
    var mime = imageMimeMap[ext] || 'image/png';
    var uri  = getFileAsDataUri(fileName, mime);
    if (uri) images[fileName] = uri;
  });

  return { fonts: fonts, images: images };
}

// ============================================================================
// Getters — Page Settings
// ============================================================================

/**
 * Returns current page layout settings.
 * @returns {Object}
 */
function getPageSettings(allProps, preset) {
  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);
  return {
    sizePreset:       p[PROP_KEYS.PAGE_SIZE_PRESET]                || pr.page.sizePreset,
    width:            parseFloat(p[PROP_KEYS.PAGE_WIDTH])          || pr.page.width,
    height:           parseFloat(p[PROP_KEYS.PAGE_HEIGHT])         || pr.page.height,
    marginTop:        parseFloat(p[PROP_KEYS.PAGE_MARGIN_TOP])     || pr.page.marginTop,
    marginBottom:     parseFloat(p[PROP_KEYS.PAGE_MARGIN_BOTTOM])  || pr.page.marginBottom,
    marginInner:      parseFloat(p[PROP_KEYS.PAGE_MARGIN_INNER])   || pr.page.marginInner,
    marginOuter:      parseFloat(p[PROP_KEYS.PAGE_MARGIN_OUTER])   || pr.page.marginOuter,
    pageBuffer:       parseInt(p[PROP_KEYS.PAGE_BUFFER])           || pr.page.pageBuffer,
    frontMatterPages: parseInt(p[PROP_KEYS.FRONT_MATTER_PAGES])    || pr.frontMatterPages
  };
}

// ============================================================================
// Getters — Footer Settings
// ============================================================================

/**
 * Returns current footer / running label settings.
 * @returns {Object}
 */
function getFooterSettings(allProps, preset) {
  var p         = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr        = preset   || getActiveBrandPreset(p);
  var showLabel = p[PROP_KEYS.SHOW_RUNNING_LABEL];
  return {
    pageNumberPosition:   p[PROP_KEYS.FOOTER_PAGE_NUMBER_POSITION] || pr.footer.pageNumberPosition,
    footerRule:           p[PROP_KEYS.FOOTER_RULE]                 || pr.footer.footerRule,
    showRunningLabel:     showLabel != null ? showLabel === 'true'  : pr.footer.showRunningLabel,
    runningLabelPosition: p[PROP_KEYS.RUNNING_LABEL_POSITION]      || pr.footer.runningLabelPosition
  };
}

// ============================================================================
// Getters — Heading Styles
// ============================================================================

/**
 * Reads an optional numeric heading property with graceful fallback.
 * Handles both null (individual getProperty result) and undefined
 * (object key miss from a getProperties batch result).
 * @private
 */
function getOptionalPt_(stored, presetVal) {
  if (stored != null) {
    var n = parseFloat(stored);
    if (!isNaN(n)) return n;
  }
  return (presetVal !== undefined) ? presetVal : null;
}

/**
 * Returns the heading style for a given type, with brand-preset fallbacks.
 * spaceBefore and spaceAfter are optional — null means the CSS generation
 * falls back to the size-proportional formula for backwards compatibility.
 * @param {number} type  1–4
 * @returns {Object|null}
 */
function getHeadingStyle(type, allProps, preset) {
  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);
  var d  = pr.headingStyles[type];
  if (!d) return null;

  return {
    title: {
      font:        p[headingKey_(type, 'TITLE', 'FONT')]                        || d.title.font,
      size:        parseInt(p[headingKey_(type, 'TITLE', 'SIZE')])              || d.title.size,
      color:       p[headingKey_(type, 'TITLE', 'COLOR')]                       || d.title.color,
      align:       p[headingKey_(type, 'TITLE', 'ALIGN')]                       || d.title.align,
      weight:      p[headingKey_(type, 'TITLE', 'WEIGHT')]                      || d.title.weight,
      transform:   p[headingKey_(type, 'TITLE', 'TRANSFORM')]                   || d.title.transform,
      spacing:     parseFloat(p[headingKey_(type, 'TITLE', 'SPACING')])         || d.title.spacing,
      underline:   p[headingKey_(type, 'TITLE', 'UNDERLINE')]                   || d.title.underline,
      variant:     p[headingKey_(type, 'TITLE', 'VARIANT')]                     || d.title.variant || 'normal',
      spaceBefore: getOptionalPt_(p[headingKey_(type, 'TITLE', 'SPACEBEFORE')], d.title.spaceBefore),
      spaceAfter:  getOptionalPt_(p[headingKey_(type, 'TITLE', 'SPACEAFTER')],  d.title.spaceAfter)
    },
    subtext: {
      font:     p[headingKey_(type, 'SUB', 'FONT')]                             || d.subtext.font,
      size:     parseInt(p[headingKey_(type, 'SUB', 'SIZE')])                   || d.subtext.size,
      color:    p[headingKey_(type, 'SUB', 'COLOR')]                            || d.subtext.color,
      weight:   p[headingKey_(type, 'SUB', 'WEIGHT')]                           || d.subtext.weight,
      position: p[headingKey_(type, 'SUB', 'POSITION')]                         || d.subtext.position
    }
  };
}

/**
 * Returns heading styles for all types (1–MAX_HEADING_TYPE).
 * @returns {Object.<number, Object>}
 */
function getAllHeadingStyles(allProps, preset) {
  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);
  var styles = {};
  for (var t = 1; t <= MAX_HEADING_TYPE; t++) {
    styles[t] = getHeadingStyle(t, p, pr);
  }
  return styles;
}

// ============================================================================
// Getters — Wine Entry, Images, Colors, Welcome
// ============================================================================

function getWineEntryStyle(allProps, preset) {
  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);
  return {
    font:   p[PROP_KEYS.WINE_FONT]           || pr.wineEntry.font,
    size:   parseInt(p[PROP_KEYS.WINE_SIZE]) || pr.wineEntry.size,
    color:  p[PROP_KEYS.WINE_COLOR]          || pr.wineEntry.color,
    weight: p[PROP_KEYS.WINE_WEIGHT]         || pr.wineEntry.weight,
    style:  p[PROP_KEYS.WINE_STYLE]          || pr.wineEntry.style
  };
}

function getImageSettings(allProps, preset) {
  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);
  return {
    logo:   p[PROP_KEYS.IMAGE_LOGO]   || pr.images.logo,
    footer: p[PROP_KEYS.IMAGE_FOOTER] || pr.images.footer
  };
}

function getColorSettings(allProps, preset) {
  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);
  return {
    primary: p[PROP_KEYS.COLOR_PRIMARY] || pr.colors.primary,
    text:    p[PROP_KEYS.COLOR_TEXT]    || pr.colors.text
  };
}

function getWelcomeSettings(allProps, preset) {
  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);

  var showTitle = p[PROP_KEYS.SHOW_TITLE_PAGE];
  var showDate  = p[PROP_KEYS.SHOW_DATE];
  var line1     = p[PROP_KEYS.WELCOME_LINE1];
  var line2     = p[PROP_KEYS.WELCOME_LINE2];
  var line3     = p[PROP_KEYS.WELCOME_LINE3];

  return {
    showTitlePage: showTitle != null ? showTitle === 'true' : pr.welcome.showTitlePage,
    showDate:      showDate  != null ? showDate  === 'true' : pr.welcome.showDate,
    line1:         line1     != null ? line1     : pr.welcome.line1,
    line2:         line2     != null ? line2     : pr.welcome.line2,
    line3:         line3     != null ? line3     : pr.welcome.line3
  };
}

// ============================================================================
// Aggregate Getter — called once on Settings dialog open
// ============================================================================

/**
 * Returns all settings needed to fully populate the Settings dialog UI.
 * fontPool and imagePool are the filenames available in the Drive folder.
 * Asset binary data (for live previews) is loaded separately via getAssetPreviewData().
 * @returns {Object}
 */
function getAllSettings() {
  var allProps = PropertiesService.getDocumentProperties().getProperties();
  var preset   = getActiveBrandPreset(allProps);
  var assets   = scanDriveFolder_();

  return {
    brandName:     allProps[PROP_KEYS.BRAND_NAME] || 'BYGONE',
    brandPresets:  getBrandPresetNames(),
    pageSizes:     getPageSizePresetNames(),
    pageSettings:  getPageSettings(allProps, preset),
    headingStyles: getAllHeadingStyles(allProps, preset),
    wineEntry:     getWineEntryStyle(allProps, preset),
    images:        getImageSettings(allProps, preset),
    colors:        getColorSettings(allProps, preset),
    welcome:       getWelcomeSettings(allProps, preset),
    footer:        getFooterSettings(allProps, preset),
    fontPool:      assets.fonts,
    imagePool:     assets.images
  };
}

// ============================================================================
// Setters — Heading Styles
// ============================================================================

function saveHeadingStyleProp(type, group, prop, value) {
  try {
    PropertiesService.getDocumentProperties().setProperty(headingKey_(type, group, prop), value.toString());
    return { success: true, message: 'Saved.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveHeadingStyle(type, styleObj) {
  try {
    var props = {};
    if (styleObj.title) {
      Object.keys(styleObj.title).forEach(function(k) {
        props[headingKey_(type, 'TITLE', k.toUpperCase())] = styleObj.title[k].toString();
      });
    }
    if (styleObj.subtext) {
      Object.keys(styleObj.subtext).forEach(function(k) {
        props[headingKey_(type, 'SUB', k.toUpperCase())] = styleObj.subtext[k].toString();
      });
    }
    PropertiesService.getDocumentProperties().setProperties(props);
    return { success: true, message: 'Heading style saved for Type ' + type + '.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// Setters — Wine Entry, Page, Footer
// ============================================================================

function saveWineEntryStyle(styleObj) {
  try {
    var p = PropertiesService.getDocumentProperties();
    if (styleObj.font)   p.setProperty(PROP_KEYS.WINE_FONT,   styleObj.font);
    if (styleObj.size)   p.setProperty(PROP_KEYS.WINE_SIZE,   styleObj.size.toString());
    if (styleObj.color)  p.setProperty(PROP_KEYS.WINE_COLOR,  styleObj.color);
    if (styleObj.weight) p.setProperty(PROP_KEYS.WINE_WEIGHT, styleObj.weight);
    if (styleObj.style)  p.setProperty(PROP_KEYS.WINE_STYLE,  styleObj.style);
    return { success: true, message: 'Wine entry style saved.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Saves page layout settings.
 * @param {Object} pageObj
 * @returns {Object}
 */
function savePageSettings(pageObj) {
  try {
    var p     = PropertiesService.getDocumentProperties();
    var props = {};
    if (pageObj.sizePreset !== undefined)       props[PROP_KEYS.PAGE_SIZE_PRESET]   = pageObj.sizePreset;
    if (pageObj.width !== undefined)            props[PROP_KEYS.PAGE_WIDTH]         = pageObj.width.toString();
    if (pageObj.height !== undefined)           props[PROP_KEYS.PAGE_HEIGHT]        = pageObj.height.toString();
    if (pageObj.marginTop !== undefined)        props[PROP_KEYS.PAGE_MARGIN_TOP]    = pageObj.marginTop.toString();
    if (pageObj.marginBottom !== undefined)     props[PROP_KEYS.PAGE_MARGIN_BOTTOM] = pageObj.marginBottom.toString();
    if (pageObj.marginInner !== undefined)      props[PROP_KEYS.PAGE_MARGIN_INNER]  = pageObj.marginInner.toString();
    if (pageObj.marginOuter !== undefined)      props[PROP_KEYS.PAGE_MARGIN_OUTER]  = pageObj.marginOuter.toString();
    if (pageObj.pageBuffer !== undefined)       props[PROP_KEYS.PAGE_BUFFER]        = pageObj.pageBuffer.toString();
    if (pageObj.frontMatterPages !== undefined) props[PROP_KEYS.FRONT_MATTER_PAGES] = pageObj.frontMatterPages.toString();
    p.setProperties(props);
    return { success: true, message: 'Page settings saved.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Saves footer / running label settings.
 * @param {Object} footerObj
 * @returns {Object}
 */
function saveFooterSettings(footerObj) {
  try {
    var p = PropertiesService.getDocumentProperties();
    if (footerObj.pageNumberPosition !== undefined) p.setProperty(PROP_KEYS.FOOTER_PAGE_NUMBER_POSITION, footerObj.pageNumberPosition);
    if (footerObj.footerRule         !== undefined) p.setProperty(PROP_KEYS.FOOTER_RULE,                 footerObj.footerRule);
    if (footerObj.showRunningLabel   !== undefined) p.setProperty(PROP_KEYS.SHOW_RUNNING_LABEL,          footerObj.showRunningLabel.toString());
    if (footerObj.runningLabelPosition !== undefined) p.setProperty(PROP_KEYS.RUNNING_LABEL_POSITION,    footerObj.runningLabelPosition);
    return { success: true, message: 'Footer settings saved.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================================
// Setters — Colors, Welcome
// ============================================================================

function saveColor(slot, value) {
  var propKey = PROP_KEYS['COLOR_' + slot.toUpperCase()];
  if (!propKey) return { success: false, message: 'Invalid color slot: ' + slot };
  if (!/^#[0-9A-Fa-f]{6}$/.test(value)) {
    return { success: false, message: 'Invalid hex color. Use format #RRGGBB.' };
  }
  PropertiesService.getDocumentProperties().setProperty(propKey, value);
  return { success: true, message: 'Color saved.' };
}

function saveWelcomeText(line1, line2, line3) {
  var p = PropertiesService.getDocumentProperties();
  p.setProperty(PROP_KEYS.WELCOME_LINE1, line1 || '');
  p.setProperty(PROP_KEYS.WELCOME_LINE2, line2 || '');
  p.setProperty(PROP_KEYS.WELCOME_LINE3, line3 || '');
  return { success: true, message: 'Welcome text saved.' };
}

function saveWelcomeToggles(showTitlePage, showDate) {
  var p = PropertiesService.getDocumentProperties();
  p.setProperty(PROP_KEYS.SHOW_TITLE_PAGE, showTitlePage.toString());
  p.setProperty(PROP_KEYS.SHOW_DATE,       showDate.toString());
  return { success: true, message: 'Title page settings saved.' };
}

// ============================================================================
// Setters — Asset Files (Fonts & Images)
// ============================================================================

/**
 * Uploads a font file to the spreadsheet's Drive folder.
 * Returns the updated font pool so the dialog can refresh immediately.
 *
 * @param {string} base64Data  Base64-encoded file content.
 * @param {string} fileName    Original filename including extension.
 * @returns {Object} { success, message, fileName?, fontPool? }
 */
function saveFontFile(base64Data, fileName) {
  var ext = extensionOf(fileName);
  if (['otf', 'ttf', 'woff', 'woff2'].indexOf(ext) === -1) {
    return { success: false, message: 'Unsupported font format. Use .otf, .ttf, .woff, or .woff2.' };
  }
  var result = saveFileToDrive_(base64Data, fileName, getMimeForExtension(ext));
  if (result.success) {
    result.fontPool = getAvailableFonts();
  }
  return result;
}

/**
 * Uploads an image file to the spreadsheet's Drive folder.
 * Does NOT assign the image to any slot — use saveImageAssignment() for that.
 * Returns the updated image pool so the dialog can refresh immediately.
 *
 * @param {string} base64Data  Base64-encoded file content.
 * @param {string} fileName    Original filename including extension.
 * @returns {Object} { success, message, fileName?, imagePool? }
 */
function uploadImageFile(base64Data, fileName) {
  var ext = extensionOf(fileName);
  if (['png', 'jpg', 'jpeg', 'svg', 'webp'].indexOf(ext) === -1) {
    return { success: false, message: 'Unsupported image format. Use .png, .jpg, .svg, or .webp.' };
  }
  var mimeMap = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
    svg: 'image/svg+xml', webp: 'image/webp'
  };
  var result = saveFileToDrive_(base64Data, fileName, mimeMap[ext] || 'application/octet-stream');
  if (result.success) {
    result.imagePool = getAvailableImages();
  }
  return result;
}

/**
 * Assigns an image file (already in Drive) to a slot (logo or footer).
 * Stores the filename in DocumentProperties. No Drive I/O.
 *
 * @param {string} slot      'logo' or 'footer'
 * @param {string} fileName  Filename of the image in the Drive folder.
 * @returns {Object} { success, message }
 */
function saveImageAssignment(slot, fileName) {
  var propKey = PROP_KEYS['IMAGE_' + slot.toUpperCase()];
  if (!propKey) return { success: false, message: 'Invalid image slot: ' + slot };
  if (!fileName || !fileName.trim()) return { success: false, message: 'Filename cannot be empty.' };
  PropertiesService.getDocumentProperties().setProperty(propKey, fileName.trim());
  var label = slot.charAt(0).toUpperCase() + slot.slice(1);
  return { success: true, message: label + ' image set to "' + fileName + '".' };
}

/**
 * Resets all settings to the active brand's defaults.
 * @returns {Object}
 */
function resetAllDefaults() {
  var brandName = getBrandName();
  return loadBrandPreset(brandName);
}

// ============================================================================
// Internal Helpers
// ============================================================================

/**
 * Saves a file to the spreadsheet's Drive folder, replacing any existing file
 * with the same name.
 * @private
 */
function saveFileToDrive_(base64Data, fileName, mimeType) {
  try {
    var ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var parents = ssFile.getParents();
    if (!parents.hasNext()) {
      return { success: false, message: 'Cannot find spreadsheet folder.' };
    }
    var folder = parents.next();

    // Replace existing file with the same name
    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    var decoded = Utilities.base64Decode(base64Data);
    var blob    = Utilities.newBlob(decoded, mimeType, fileName);
    folder.createFile(blob);

    return { success: true, message: 'File saved: ' + fileName, fileName: fileName };
  } catch (e) {
    return { success: false, message: 'Upload error: ' + e.toString() };
  }
}

function extensionOf(fileName) {
  return (fileName || '').split('.').pop().toLowerCase();
}

function getMimeForExtension(ext) {
  var map = { otf: 'font/otf', ttf: 'font/ttf', woff: 'font/woff', woff2: 'font/woff2' };
  return map[ext] || 'application/octet-stream';
}

function getFontFormatForExtension(ext) {
  var map = { otf: 'opentype', ttf: 'truetype', woff: 'woff', woff2: 'woff2' };
  return map[ext] || 'opentype';
}
