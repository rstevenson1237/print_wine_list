// config.gs

/**
 * @fileoverview Configuration constants for the Wine List Generator.
 * Centralizes all magic numbers, settings, and mappings.
 *
 * Brand presets define per-property visual configurations.
 * User-facing settings can be overridden via the Settings dialog.
 */

/**
 * Sheet names used throughout the application.
 * @enum {string}
 */
const SHEETS = {
  LIST: 'List',
  DATA: 'Data',
  SECTIONS: 'Sections'
};

/**
 * Section heading types.
 */
const SECTION_TYPES = {
  1: 'STYLE',
  2: 'COUNTRY',
  3: 'REGION',
  4: 'APPELLATION'
};

const MAX_HEADING_TYPE = 4;

// ============================================================================
// Page Size Presets
// ============================================================================

/**
 * Named page size presets. Width × Height in inches.
 * Key is the display label; value is { width, height }.
 */
const PAGE_SIZE_PRESETS = {
  'Half Letter (5.5 × 8.5)':  { width: 5.5,  height: 8.5  },
  'Small Book (6 × 8.5)':     { width: 6,    height: 8.5  },
  'Bygone (8 × 13)':          { width: 8,    height: 13   },
  'Ruxton (8 × 13)':          { width: 8,    height: 13   },
  'US Letter (8.5 × 11)':     { width: 8.5,  height: 11   },
  'Large Book (9 × 13)':      { width: 9,    height: 13   }
};

/**
 * Returns an array of page size preset names for UI dropdowns.
 */
function getPageSizePresetNames() {
  return Object.keys(PAGE_SIZE_PRESETS);
}

// ============================================================================
// Brand Presets
// ============================================================================

/**
 * Complete brand configurations per property.
 * Each preset contains every configurable setting so that loading a preset
 * gives a fully functional wine list without additional tweaks.
 */
const BRAND_PRESETS = {

  BYGONE: {
    label: 'The Bygone',
    page: {
      sizePreset: 'Bygone (8 × 13)',
      width: 8,
      height: 13,
      marginTop:    0.75,
      marginBottom: 0.5,
      marginInner:  0.5,
      marginOuter:  1.0,
      pageBuffer: 100
    },
    colors: {
      primary: '#B88800',
      text: '#333333'
    },
    images: {
      logo: 'Bygone.png',
      footer: 'Bygone_Footer.png'
    },
    welcome: {
      showTitlePage: true,
      showDate: true,
      line1: 'We welcome you to The Bygone',
      line2: 'We have carefully and meticulously selected over 700 wines to complement the cuisine crafted by our executive chef. By providing an extensive selection of French wines, we hope to enhance your experience here at the Bygone.',
      line3: 'We are sure you will find a bottle from our collection that will please your discerning palate, and elevate your dining experience. Please allow our sommelier to help you navigate our wine list and select a special and rare bottle of wine for your occasion; Or simply take a look at our beautiful collection of wines available by the glass. Enjoy!'
    },
    headingStyles: {
      1: {
        title:   { font: 'Cormier-Regular.otf', size: 20, color: '#B88800', align: 'center', weight: 'normal', transform: 'uppercase', spacing: 2, underline: 'none', variant: 'normal' },
        subtext: { font: 'ApexNew-Book.otf',     size: 14, color: '#333333', weight: 'normal', position: 'below' }
      },
      2: {
        title:   { font: 'Minerva Regular.otf', size: 22, color: '#B88800', align: 'left', weight: 'normal', transform: 'none', spacing: 0, underline: 'none', variant: 'normal' },
        subtext: { font: 'ApexNew-Book.otf',     size: 12, color: '#333333', weight: 'normal', position: 'below' }
      },
      3: {
        title:   { font: 'ApexSansBoldST.ttf', size: 14, color: '#555555', align: 'left', weight: 'bold', transform: 'none', spacing: 1, underline: 'partial', variant: 'small-caps' },
        subtext: { font: 'ApexNew-Book.otf',    size: 11, color: '#333333', weight: 'normal', position: 'inline' }
      },
      4: {
        title:   { font: 'ApexSansBoldST.ttf', size: 9, color: '#000000', align: 'left', weight: 'bold', transform: 'uppercase', spacing: 5, underline: 'text', variant: 'normal' },
        subtext: { font: 'ApexNew-Book.otf',    size: 10, color: '#333333', weight: 'normal', position: 'inline' }
      }
    },
    wineEntry: {
      font: 'ApexNew-Book.otf', size: 10, color: '#333333', weight: 'normal', style: 'normal'
    },
    footer: {
      pageNumberPosition: 'center',
      footerRule:         'none',
      showRunningLabel:   true,
      runningLabelPosition: 'header'
    },
    frontMatterPages: 2
  },

  RUXTON: {
    label: 'Ruxton',
    page: {
      sizePreset: 'Ruxton (8 × 13)',  // item 2
      width: 8,                        // item 2
      height: 13,                      // item 2
      marginTop:    0.5,               // item 4 (was 0.6)
      marginBottom: 0.5,               // item 4 (was 0.75)
      marginInner:  1.0,               // unchanged
      marginOuter:  0.5,               // item 4 (was 0.75)
      pageBuffer: 80
    },
    colors: {
      primary: '#2D4A1E',              // item 1 (was #3B5F3B)
      text: '#222222'
    },
    images: {
      logo: 'Ruxton.png',
      footer: 'Ruxton_Footer.png'
    },
    welcome: {
      showTitlePage: true,
      showDate: true,
      line1: '',
      line2: '',
      line3: ''
    },
    headingStyles: {
      1: {
        // Page Header — script display font, brand green, centered, generous spacing
        title:   { font: 'Ruxton-Heading.otf', size: 39, color: '#2D4A1E', align: 'center', weight: 'normal', transform: 'none', spacing: 1, underline: 'none', variant: 'normal', spaceBefore: 50, spaceAfter: 32 },
        // items 1 (color), 3 (size), 9 (spaceBefore/spaceAfter)
        subtext: { font: 'Ruxton-Body.otf', size: 14, color: '#222222', weight: 'normal', position: 'below' }
      },
      2: {
        // Country / Wine Color — display font, centered, all-caps
        title:   { font: 'Ruxton-Heading.otf', size: 22, color: '#2D4A1E', align: 'center', weight: 'normal', transform: 'uppercase', spacing: 1, underline: 'none', variant: 'normal', spaceBefore: 28, spaceAfter: 9 },
        // items 1 (color), 5 (align), 6 (size), 9 (spaceBefore/spaceAfter)
        subtext: { font: 'Ruxton-Body.otf', size: 11, color: '#555555', weight: 'normal', position: 'below' }
      },
      3: {
        // Wine Type category label — brand green, wide tracking, generous space above
        title:   { font: 'Ruxton-Body.otf', size: 13, color: '#2D4A1E', align: 'left', weight: 'bold', transform: 'uppercase', spacing: 0.5, underline: 'none', variant: 'normal', spaceBefore: 36, spaceAfter: 0 },
        // items 1 (color), 7 (underline none), 9 (spaceBefore/spaceAfter)
        subtext: { font: 'Ruxton-Body.otf', size: 10, color: '#555555', weight: 'normal', position: 'below' }
      },
      4: {
        // Specific type / appellation — body font, subtle
        title:   { font: 'Ruxton-Body.otf', size: 11, color: '#222222', align: 'left', weight: 'bold', transform: 'none', spacing: 0, underline: 'text', variant: 'normal', spaceBefore: 14, spaceAfter: 5 },
        // item 9 (spaceBefore/spaceAfter)
        subtext: { font: 'Ruxton-Body.otf', size: 10, color: '#555555', weight: 'normal', position: 'inline' }
      }
    },
    wineEntry: {
      font: 'Ruxton-Body.otf', size: 12, color: '#222222', weight: 'normal', style: 'italic'
    },
    footer: {
      pageNumberPosition: 'outer',
      footerRule:         'double',
      showRunningLabel:   true,
      runningLabelPosition: 'header'
    },
    frontMatterPages: 2
  }

};

/**
 * Returns an array of brand preset names for UI dropdowns.
 */
function getBrandPresetNames() {
  return Object.keys(BRAND_PRESETS);
}

// ============================================================================
// Dynamic Page Config
// ============================================================================

/**
 * Returns the active page configuration, reading from DocumentProperties
 * with fallback to the active brand preset or hard-coded defaults.
 *
 * @param {Object} [allProps]  Optional pre-fetched DocumentProperties object.
 * @param {Object} [preset]    Optional pre-resolved brand preset.
 * @returns {Object} Page configuration with computed helper methods.
 */
function getPageConfig(allProps, preset) {

  var p  = allProps || PropertiesService.getDocumentProperties().getProperties();
  var pr = preset   || getActiveBrandPreset(p);

  var config = {
    WIDTH:         parseFloat(p[PROP_KEYS.PAGE_WIDTH])         || pr.page.width,
    HEIGHT:        parseFloat(p[PROP_KEYS.PAGE_HEIGHT])        || pr.page.height,
    MARGIN_TOP:    parseFloat(p[PROP_KEYS.PAGE_MARGIN_TOP])    || pr.page.marginTop,
    MARGIN_BOTTOM: parseFloat(p[PROP_KEYS.PAGE_MARGIN_BOTTOM]) || pr.page.marginBottom,
    MARGIN_INNER:  parseFloat(p[PROP_KEYS.PAGE_MARGIN_INNER])  || pr.page.marginInner,
    MARGIN_OUTER:  parseFloat(p[PROP_KEYS.PAGE_MARGIN_OUTER])  || pr.page.marginOuter,
    PAGE_BUFFER:   parseInt(p[PROP_KEYS.PAGE_BUFFER])          || pr.page.pageBuffer,
    PPI: 72,
    FRONT_MATTER_PAGES: parseInt(p[PROP_KEYS.FRONT_MATTER_PAGES]) || pr.frontMatterPages
  };

  config.getUsableHeightPts = function() {
    return (this.HEIGHT - this.MARGIN_TOP - this.MARGIN_BOTTOM) * this.PPI;
  };
  config.getUsableWidthPts = function() {
    return (this.WIDTH - this.MARGIN_INNER - this.MARGIN_OUTER) * this.PPI;
  };

  return config;
}

// ============================================================================
// Fixed Element Heights (used by pagination)
// ============================================================================

const ELEMENT_HEIGHTS = {
  // Wine entry layout (used by both html.gs and pagination.gs)
  WINE_ENTRY_MARGIN:  3,    // px, top and bottom each
  WINE_ENTRY_PADDING: 1,    // px, top and bottom each
  LINE_HEIGHT:        1.3,  // body line-height multiplier

  // Heading layout
  SUBTEXT_LINE: 16,
  MIN_WINES_PER_SPLIT: 3,
  RUXTON_FOOTER_ICON_SIZE: 18
};

/**
 * Computes the effective rendered height of a single wine entry in points.
 * Matches the CSS: margin-top + padding-top + (fontSize × lineHeight) + padding-bottom + margin-bottom
 *
 * @param {Object} wineEntry  Resolved wine entry settings (needs .size at minimum).
 * @returns {number} Height in px/pt.
 */
function getWineEntryHeight(wineEntry) {
  var m = ELEMENT_HEIGHTS.WINE_ENTRY_MARGIN;
  var p = ELEMENT_HEIGHTS.WINE_ENTRY_PADDING;
  return (m + p) * 2 + (wineEntry.size * ELEMENT_HEIGHTS.LINE_HEIGHT);
}

// ============================================================================
// R365 / Data Import Mappings (unchanged)
// ============================================================================

const R365_HEADER_MAP = {
  'ItemName': 'Item',
  'ItemCategory1': 'Category1',
  'ItemCategory2': 'Category2',
  'ItemCategory3': 'Category3',
  'VendorName': 'Vendor',
  'TransactionDate': 'LastPurchase',
  'PurchaseUnit': 'UofM',
  'AmountEach': 'Cost'
};

const VARIANCE_HEADER_MAP = {
  'Item Name': 'Item',
  'Unit of Measure': 'UofM',
  'Current $ Cost': 'Cost',
  'Current Qty': 'Count',
  'Item Category 1': 'Category1',
  'Item Category 2': 'Category2',
  'Item Category 3': 'Category3'
};

const TOAST_PRICING_HEADER_MAP = {
  'Name':       'Name',
  'Base Price': 'BasePrice',
  'Archived':   'Archived',
  'Modifier':   'Modifier'
};

const DATA_SHEET_HEADERS = [
  'Item', 'Name', 'Section', 'Vintage', 'MenuPrice', 'Bin', 'Cost',
  'UofM', 'StorageLocation', 'NewPrice', 'ToastPrice', 'Count',
  'Vendor', 'Par', 'LastPurchase', 'Order', 'Notes', 'Category1',
  'Category2', 'Category3'
];

// Legacy — kept for any old references
const BRAND = {};

// Legacy shim — old code that references PAGE_CONFIG directly will still work
// but new code should use getPageConfig()
const PAGE_CONFIG = {
  WIDTH: 8, HEIGHT: 13,
  MARGIN_TOP: 0.75, MARGIN_BOTTOM: 1.0,
  MARGIN_INNER: 1.25, MARGIN_OUTER: 0.75,
  PPI: 72, FRONT_MATTER_PAGES: 2,
  getUsableHeightPts: function() { return (this.HEIGHT - this.MARGIN_TOP - this.MARGIN_BOTTOM) * this.PPI; },
  getUsableWidthPts: function() { return (this.WIDTH - this.MARGIN_INNER - this.MARGIN_OUTER) * this.PPI; }
};
