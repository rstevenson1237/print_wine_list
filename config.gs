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
      marginTop:    0.75,  // unchanged — matches InDesign A-Master
      marginBottom: 0.5,   // was 1.0  — InDesign uses 0.5"
      marginInner:  0.5,   // was 1.25 — InDesign inner is the tight gutter side
      marginOuter:  1.0,   // was 0.75 — InDesign outer is the generous thumb margin
      pageBuffer: 100
    },
    colors: {
      primary: '#B88800',  // was #B88500 — closer CMYK-to-sRGB conversion of C10 M35 Y100 K20
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
        // Wine Type / Style — Cormier display font, gold, centered
        title:   { font: 'Cormier-Regular.otf', size: 20, color: '#B88800', align: 'center', weight: 'normal', transform: 'uppercase', spacing: 2, underline: 'none', variant: 'normal' },
        subtext: { font: 'ApexNew-Book.otf',     size: 14, color: '#333333', weight: 'normal', position: 'below' }
      },
      2: {
        // Country — Minerva Regular, gold, left-aligned
        title:   { font: 'Minerva Regular.otf', size: 22, color: '#B88800', align: 'left', weight: 'normal', transform: 'none', spacing: 0, underline: 'none', variant: 'normal' },
        subtext: { font: 'ApexNew-Book.otf',     size: 12, color: '#333333', weight: 'normal', position: 'below' }
      },
      3: {
        // Region — Apex Sans, dark gray, small-caps
        title:   { font: 'ApexSansBoldST.ttf', size: 14, color: '#555555', align: 'left', weight: 'bold', transform: 'none', spacing: 1, underline: 'partial', variant: 'small-caps' },
        subtext: { font: 'ApexNew-Book.otf',    size: 11, color: '#333333', weight: 'normal', position: 'inline' }
      },
      4: {
        // Appellation — Apex Sans, all-caps, wide tracking, smaller
        title:   { font: 'ApexSansBoldST.ttf', size: 9, color: '#000000', align: 'left', weight: 'bold', transform: 'uppercase', spacing: 5, underline: 'text', variant: 'normal' },
        subtext: { font: 'ApexNew-Book.otf',    size: 10, color: '#333333', weight: 'normal', position: 'inline' }
      }
    },
    wineEntry: {
      font: 'ApexNew-Book.otf', size: 10, color: '#333333', weight: 'normal', style: 'normal'  // was size: 13
    },
    footer: {
      style: 'image',
      showRunningLabel: true,
      runningLabelPosition: 'header'
    },
    frontMatterPages: 2
  },

  RUXTON: {
    label: 'Ruxton',
    page: {
      sizePreset: 'US Letter (8.5 × 11)',
      width: 8.5,
      height: 11,
      marginTop: 0.6,
      marginBottom: 0.75,
      marginInner: 1.0,
      marginOuter: 0.75,
      pageBuffer: 80
    },
    colors: {
      primary: '#3B5F3B',
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
        title:   { font: 'Ruxton-Heading.otf', size: 32, color: '#3B5F3B', align: 'center', weight: 'normal', transform: 'none', spacing: 1, underline: 'none' },
        subtext: { font: 'Ruxton-Body.otf',    size: 14, color: '#222222', weight: 'normal', position: 'below' }
      },
      2: {
        title:   { font: 'Ruxton-Heading.otf', size: 16, color: '#3B5F3B', align: 'left', weight: 'bold', transform: 'uppercase', spacing: 1, underline: 'none' },
        subtext: { font: 'Ruxton-Body.otf',    size: 11, color: '#555555', weight: 'normal', position: 'below' }
      },
      3: {
        title:   { font: 'Ruxton-Body.otf', size: 13, color: '#3B5F3B', align: 'left', weight: 'bold', transform: 'uppercase', spacing: 0.5, underline: 'full' },
        subtext: { font: 'Ruxton-Body.otf', size: 10, color: '#555555', weight: 'normal', position: 'below' }
      },
      4: {
        title:   { font: 'Ruxton-Body.otf', size: 11, color: '#222222', align: 'left', weight: 'bold', transform: 'none', spacing: 0, underline: 'text' },
        subtext: { font: 'Ruxton-Body.otf', size: 10, color: '#555555', weight: 'normal', position: 'inline' }
      }
    },
    wineEntry: {
      font: 'Ruxton-Body.otf', size: 12, color: '#222222', weight: 'normal', style: 'italic'
    },
    footer: {
      style: 'ruxton',
      showRunningLabel: true,
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
 * Returns the active page configuration, reading from stored settings
 * with fallback to the active brand preset or hard-coded defaults.
 *
 * @returns {Object} Page configuration with computed helper methods.
 */
function getPageConfig() {
  var p = PropertiesService.getUserProperties();
  var brandName = p.getProperty('BRAND_NAME') || 'BYGONE';
  var preset = BRAND_PRESETS[brandName] || BRAND_PRESETS.BYGONE;
  var defaults = preset.page;

  var config = {
    WIDTH:         parseFloat(p.getProperty('PAGE_WIDTH'))         || defaults.width,
    HEIGHT:        parseFloat(p.getProperty('PAGE_HEIGHT'))        || defaults.height,
    MARGIN_TOP:    parseFloat(p.getProperty('PAGE_MARGIN_TOP'))    || defaults.marginTop,
    MARGIN_BOTTOM: parseFloat(p.getProperty('PAGE_MARGIN_BOTTOM')) || defaults.marginBottom,
    MARGIN_INNER:  parseFloat(p.getProperty('PAGE_MARGIN_INNER'))  || defaults.marginInner,
    MARGIN_OUTER:  parseFloat(p.getProperty('PAGE_MARGIN_OUTER'))  || defaults.marginOuter,
    PAGE_BUFFER:   parseInt(p.getProperty('PAGE_BUFFER'))          || defaults.pageBuffer,
    PPI: 72,
    FRONT_MATTER_PAGES: parseInt(p.getProperty('FRONT_MATTER_PAGES')) || defaults.frontMatterPages
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
  WINE_ENTRY: 20,
  SUBTEXT_LINE: 16,
  MIN_WINES_PER_SPLIT: 3,
  HEADING_VERTICAL_PADDING: 20
};

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
