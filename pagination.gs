// pagination.gs

/**
 * @fileoverview Heuristic pagination calculator for the wine list.
 *
 * Now uses dynamic page config passed as a parameter rather than
 * the global PAGE_CONFIG constant.
 */

// ============================================================================
// Public API
// ============================================================================

/**
 * Calculates pagination for the wine list.
 *
 * @param {Array<Object>} sections Ordered array from getSectionData().
 * @param {Map<number, Array<Object>>} wineMap From buildWineMap().
 * @param {Object} headingStyles From getAllHeadingStyles().
 * @param {Object} pageConfig From getPageConfig().
 * @returns {Object}
 */
function calculatePagination(sections, wineMap, headingStyles, pageConfig, wineEntry, footerSettings) {
  var usableHeight      = pageConfig.getUsableHeightPts() - pageConfig.PAGE_BUFFER;
  var frontMatterOffset = pageConfig.FRONT_MATTER_PAGES;
  var wineEntryHeight   = getWineEntryHeight(wineEntry);

  // Running label (header position) consumes space at the top of every page after the first.
  var runningLabelHeight = (footerSettings && footerSettings.showRunningLabel &&
                            footerSettings.runningLabelPosition === 'header')
    ? ELEMENT_HEIGHTS.RUNNING_LABEL_HEADER : 0;

  var currentHeight = 0;
  var currentPage   = 1;

  var pageBreaks     = new Set();
  var winePageBreaks = new Map();
  var tocPageNumbers = new Map();

  // Cache for estimateHeadingHeight — at most 8 unique values (4 types × 2 subtext states).
  var headingHeightCache = {};

  for (var s = 0; s < sections.length; s++) {
    var section    = sections[s];
    var typeLevel  = section.type;

    var headingHeight = estimateHeadingHeight(typeLevel, section.subtext, headingStyles, headingHeightCache);
    var wines         = (section.code > 0 && wineMap.has(section.code)) ? wineMap.get(section.code) : [];
    var wineHeight    = wines.length * wineEntryHeight;
    var totalHeight   = headingHeight + wineHeight;

    // Rule 1: Type 1 or explicit forceNewPage starts a new page
    var isForced = (typeLevel === 1) || (section.forceNewPage === true);
    if (isForced && (currentHeight > 0 || currentPage > 1)) {
      pageBreaks.add(s);
      currentPage++;
      currentHeight = runningLabelHeight;
    }

    // Rule 2: Does the section fit?
    var fitsOnCurrentPage = (currentHeight + totalHeight) <= usableHeight;
    var fitsOnEmptyPage   = totalHeight <= usableHeight;

    if (fitsOnCurrentPage) {
      recordTOCEntry_(section, currentPage + frontMatterOffset, tocPageNumbers, s);
      currentHeight += totalHeight;

    } else if (currentHeight > 0 && fitsOnEmptyPage) {
      pageBreaks.add(s);
      currentPage++;
      currentHeight = runningLabelHeight;
      recordTOCEntry_(section, currentPage + frontMatterOffset, tocPageNumbers, s);
      currentHeight += totalHeight;

    } else {
      if (currentHeight > 0) {
        pageBreaks.add(s);
        currentPage++;
        currentHeight = runningLabelHeight;
      }

      recordTOCEntry_(section, currentPage + frontMatterOffset, tocPageNumbers, s);

      if (wines.length === 0) {
        currentHeight += headingHeight;
      } else {
        var splitResult = splitOversizedSection_(
          headingHeight, wines.length, usableHeight, currentPage, wineEntryHeight, runningLabelHeight
        );
        winePageBreaks.set(s, splitResult.breaks);
        currentPage   = splitResult.endPage;
        currentHeight = splitResult.endHeight;
      }
    }
  }

  Logger.log('Pagination: ' + currentPage + ' content pages, ' +
    pageBreaks.size + ' section breaks, ' + sections.length + ' sections');

  return {
    pageBreaks:     pageBreaks,
    winePageBreaks: winePageBreaks,
    tocPageNumbers: tocPageNumbers,
    totalPages:     currentPage
  };
}

// ============================================================================
// TOC Data Builder (unchanged)
// ============================================================================

function buildTOCData(sections, tocPageNumbers) {
  var tocData = [];
  for (var i = 0; i < sections.length; i++) {
    var section = sections[i];

    // Only include Types 1–3 in the TOC
    if (section.type > 3) continue;

    var key        = i + ':' + section.type + ':' + section.title;
    var pageNumber = tocPageNumbers.get(key) || '';

    tocData.push({
      type:       section.type,
      title:      section.title,
      pageNumber: pageNumber,
      slug:       createSlug('section-' + section.type + '-' + section.title)
    });
  }
  return tocData;
}

// ============================================================================
// Internal Helpers (unchanged)
// ============================================================================

function estimateHeadingHeight(typeLevel, subtext, headingStyles, cache) {
  // Memoize by (type, hasSubtext) — only 8 possible values across all sections.
  var cacheKey = typeLevel + ':' + (subtext ? '1' : '0');
  if (cache && cacheKey in cache) return cache[cacheKey];

  var style = headingStyles[typeLevel];
  if (!style) return 30;

  var ts = style.title;
  var spaceBefore = (ts.spaceBefore !== null && ts.spaceBefore !== undefined)
    ? ts.spaceBefore : Math.round(ts.size * 0.7);
  var spaceAfter  = (ts.spaceAfter  !== null && ts.spaceAfter  !== undefined)
    ? ts.spaceAfter  : Math.round(ts.size * 0.4);  // ← was 0.3, must match html.gs

  // Apply LINE_HEIGHT multiplier to heading font size — inherited from body
  var titleHeight = spaceBefore + Math.round(ts.size * ELEMENT_HEIGHTS.LINE_HEIGHT) + spaceAfter;

  var subtextHeight = 0;
  if (subtext) {
    if (style.subtext.position === 'below') {
      // Matches .subtext-type-X CSS: margin-top:2px + font-size*LINE_HEIGHT + margin-bottom:6px
      subtextHeight = 2 + Math.round(style.subtext.size * ELEMENT_HEIGHTS.LINE_HEIGHT) + 6;
    } else {
      subtextHeight = 4; // inline — no extra block height
    }
  }

  var result = titleHeight + subtextHeight;
  if (cache) cache[cacheKey] = result;
  return result;
}

/**
 * Records a section's page number in the TOC map.
 * Uses the section index as part of the key so that sections sharing the same
 * type and title (e.g., "SPAIN" appearing under multiple parent groups) each
 * receive their own correct page number rather than inheriting the first
 * occurrence's page.
 *
 * @param {Object} section       The section object.
 * @param {number} displayPage   The display page number (content page + front matter offset).
 * @param {Map}    tocMap        The TOC page-number map being built.
 * @param {number} sectionIndex  The section's index in the sections array.
 * @private
 */
function recordTOCEntry_(section, displayPage, tocMap, sectionIndex) {
  var key = sectionIndex + ':' + section.type + ':' + section.title;
  tocMap.set(key, displayPage);
}

function splitOversizedSection_(headingHeight, wineCount, usableHeight, startPage, wineEntryHeight, runningLabelHeight) {
  var breaks = new Set();
  var currentPage = startPage;
  var currentHeight = headingHeight;
  var winesOnCurrentPage = 0;
  var minWines = ELEMENT_HEIGHTS.MIN_WINES_PER_SPLIT;
  var labelHeight = runningLabelHeight || 0;

  for (var w = 0; w < wineCount; w++) {
    var wouldFit = (currentHeight + wineEntryHeight) <= usableHeight;
    var hasMinimum = winesOnCurrentPage >= minWines;

    if (!wouldFit && hasMinimum) {
      breaks.add(w);
      currentPage++;
      currentHeight = labelHeight;
      winesOnCurrentPage = 0;
    }

    currentHeight += wineEntryHeight;
    winesOnCurrentPage++;
  }

  return {
    breaks: breaks,
    endPage: currentPage,
    endHeight: currentHeight
  };
}
