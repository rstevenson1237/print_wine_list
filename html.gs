// html.gs

/**
 * @fileoverview HTML generation for the wine list.
 *
 * Changes from previous version:
 *   - Dynamic page dimensions from brand.pageConfig
 *   - Conditional title page with auto-date
 *   - Running section labels (Type 1 heading) on content pages
 *   - Footer style: 'image' (decorative) or 'rule' (thin line)
 *   - Spread-aware left/right alternating for labels and margins
 */

// ============================================================================
// Public API
// ============================================================================

/**
 * Generates the complete HTML document for the wine list.
 */
function generateHTML(sections, wineMap, pagination, assets, brand) {
  var html = generateHTMLHead(assets, brand);

  if (brand.welcome.showTitlePage) {
    html += generateTitlePage(assets.logoImageUri, brand);
  }

  html += generateTOCPage(pagination.tocData, brand);
  html += generateMainContent(sections, wineMap, pagination, brand);
  html += generateHTMLFoot();
  return html;
}

// ============================================================================
// HTML Head — Styles + Font Faces
// ============================================================================
function generateHTMLHead(assets, brand) {
  var pc   = brand.pageConfig;
  var col  = brand.colors;
  var hs   = brand.headingStyles;
  var we   = brand.wineEntry;
  var foot = brand.footer;

  // --- @font-face declarations ---
  var fontFaces = '';
  assets.fontUris.forEach(function(fontData, fileName) {
    var familyName = fileName.replace(/\.[^.]+$/, '');
    fontFaces += '\n        @font-face {\n' +
      '            font-family: \'' + familyName + '\';\n' +
      '            src: url(\'' + fontData.uri + '\') format(\'' + fontData.format + '\');\n' +
      '            font-display: swap;\n' +
      '        }';
  });

  // --- Per-type heading CSS ---
  var headingCSS = '';
  for (var t = 1; t <= MAX_HEADING_TYPE; t++) {
    var ts = hs[t].title;
    var ss = hs[t].subtext;
    var titleFont = ts.font.replace(/\.[^.]+$/, '');
    var subFont   = ss.font.replace(/\.[^.]+$/, '');

    headingCSS += '\n        .heading-type-' + t + ' {\n' +
      '            font-family: \'' + titleFont + '\', \'Georgia\', serif;\n' +
      '            font-size: ' + ts.size + 'px;\n' +
      '            color: ' + ts.color + ';\n' +
      '            text-align: ' + ts.align + ';\n' +
      '            font-weight: ' + ts.weight + ';\n' +
      '            text-transform: ' + ts.transform + ';\n' +
      '            letter-spacing: ' + ts.spacing + 'px;\n' +
      '            margin: ' + Math.round(ts.size * 0.7) + 'px 0 ' + Math.round(ts.size * 0.4) + 'px 0;\n' +
      '            page-break-after: avoid;\n' +
      '            scroll-margin-top: 20px;\n' +
      '        }';

    if (ts.underline === 'partial') {
      headingCSS += '\n        .heading-type-' + t + '::after {\n' +
        '            content: \'\';\n' +
        '            position: absolute;\n' +
        '            bottom: -2px;\n' +
        '            left: 0;\n' +
        '            width: 85%;\n' +
        '            height: 1px;\n' +
        '            background-color: ' + ts.color + ';\n' +
        '        }\n' +
        '        .heading-type-' + t + ' { position: relative; }';
    } else if (ts.underline === 'text') {
      headingCSS += '\n        .heading-type-' + t + ' {\n' +
        '            display: inline-block;\n' +
        '            border-bottom: 1px solid ' + ts.color + ';\n' +
        '        }';
    } else if (ts.underline === 'full') {
      headingCSS += '\n        .heading-type-' + t + '::after {\n' +
        '            content: \'\';\n' +
        '            position: absolute;\n' +
        '            bottom: -2px;\n' +
        '            left: 0;\n' +
        '            width: 100%;\n' +
        '            height: 1px;\n' +
        '            background-color: ' + ts.color + ';\n' +
        '        }\n' +
        '        .heading-type-' + t + ' { position: relative; }';
    }

    headingCSS += '\n        .subtext-type-' + t + ' {\n' +
      '            font-family: \'' + subFont + '\', \'Georgia\', serif;\n' +
      '            font-size: ' + ss.size + 'px;\n' +
      '            color: ' + ss.color + ';\n' +
      '            font-weight: ' + ss.weight + ';\n';
    if (ss.position === 'inline') {
      headingCSS += '            display: inline;\n' +
        '            margin-left: 8px;\n';
    } else {
      headingCSS += '            display: block;\n' +
        '            margin-top: 2px;\n' +
        '            margin-bottom: 6px;\n';
    }
    headingCSS += '        }';
  }

  var wineFont = we.font.replace(/\.[^.]+$/, '');

  // --- Footer CSS ---
  var footerCSS = '';
  if (foot.style === 'image') {
    footerCSS = '\n        @page main-pages {\n' +
      '            @bottom-center {\n' +
      '                content: counter(page);\n' +
      '                font-family: \'' + wineFont + '\', \'Georgia\', serif;\n' +
      '                font-size: 10px;\n' +
      '                color: ' + col.primary + ';\n' +
      '                background-image: url(\'' + assets.footerImageUri + '\');\n' +
      '                background-repeat: no-repeat;\n' +
      '                background-position: center center;\n' +
      '                background-size: 90% auto;\n' +
      '                height: 0.4in;\n' +
      '                padding-top: 0.15in;\n' +
      '                padding-bottom: 0.1in;\n' +
      '                display: flex;\n' +
      '                align-items: center;\n' +
      '                justify-content: center;\n' +
      '                width: 100%;\n' +
      '            }\n' +
      '        }';

  } else if (foot.style === 'ruxton') {
    // Double bar spanning full page width with page number + SVG icon on the outer edge.
    // All three @bottom-* boxes are declared on each page side so the border-top runs
    // continuously across the full content width. The outer box (right on recto pages,
    // left on verso pages) carries the page number and icon; the other two carry only
    // the bar border with empty content.
    // Note: ensure the uploaded footer file is an SVG with explicit width/height
    // attributes so Chrome renders it at the intended size in the margin box.
    var ruxtonBar  = 'border-top: 3px double ' + col.primary + '; padding-top: 5px;';
    var ruxtonText = 'font-family: \'' + wineFont + '\', \'Georgia\', serif; ' +
                     'font-size: 9px; color: ' + col.primary + '; ' +
                     'letter-spacing: 1px; vertical-align: top;';

    footerCSS =
      // Recto (right / odd) pages — number | icon on the right
      '\n        @page main-pages:right {\n' +
      '            @bottom-left   { content: ""; ' + ruxtonBar + ' }\n' +
      '            @bottom-center { content: ""; ' + ruxtonBar + ' }\n' +
      '            @bottom-right  {\n' +
      '                content: counter(page) " | " url(\'' + assets.footerImageUri + '\');\n' +
      '                ' + ruxtonBar + '\n' +
      '                ' + ruxtonText + '\n' +
      '                text-align: right;\n' +
      '            }\n' +
      '        }\n' +
      // Verso (left / even) pages — icon | number on the left
      '        @page main-pages:left {\n' +
      '            @bottom-left   {\n' +
      '                content: url(\'' + assets.footerImageUri + '\') " | " counter(page);\n' +
      '                ' + ruxtonBar + '\n' +
      '                ' + ruxtonText + '\n' +
      '                text-align: left;\n' +
      '            }\n' +
      '            @bottom-center { content: ""; ' + ruxtonBar + ' }\n' +
      '            @bottom-right  { content: ""; ' + ruxtonBar + ' }\n' +
      '        }';

  } else {
    // 'rule' — thin centred line with page number
    footerCSS = '\n        @page main-pages {\n' +
      '            @bottom-center {\n' +
      '                content: counter(page);\n' +
      '                font-family: \'' + wineFont + '\', \'Georgia\', serif;\n' +
      '                font-size: 10px;\n' +
      '                color: ' + col.primary + ';\n' +
      '                border-top: 1px solid ' + col.primary + ';\n' +
      '                padding-top: 8px;\n' +
      '                width: 100%;\n' +
      '                text-align: center;\n' +
      '            }\n' +
      '        }';
  }

  // --- Running label CSS ---
  var runningLabelCSS = '';
  if (foot.showRunningLabel) {
    if (foot.runningLabelPosition === 'header') {
      runningLabelCSS = '\n        .running-label {\n' +
        '            font-family: \'' + wineFont + '\', \'Georgia\', serif;\n' +
        '            font-size: 9px;\n' +
        '            color: ' + col.primary + ';\n' +
        '            text-transform: uppercase;\n' +
        '            letter-spacing: 2px;\n' +
        '            margin-bottom: 10px;\n' +
        '        }\n' +
        '        .running-label-left { text-align: left; }\n' +
        '        .running-label-right { text-align: right; }';
    } else {
      runningLabelCSS = '\n        .running-label {\n' +
        '            font-family: \'' + wineFont + '\', \'Georgia\', serif;\n' +
        '            font-size: 9px;\n' +
        '            color: ' + col.primary + ';\n' +
        '            text-transform: uppercase;\n' +
        '            letter-spacing: 1.5px;\n' +
        '            margin-top: 10px;\n' +
        '            padding-top: 6px;\n' +
        '        }\n' +
        '        .running-label-left { text-align: left; }\n' +
        '        .running-label-right { text-align: right; }';
    }
  }

  return '<!DOCTYPE html>\n<html lang="en">\n<head>\n' +
    '    <meta charset="UTF-8">\n' +
    '    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
    '    <title>Wine List</title>\n' +
    '    <style>\n' +
    '        /* Embedded Fonts */' + fontFaces + '\n\n' +
    '        /* Page Setup */\n' +
    '        @page {\n' +
    '            size: ' + pc.WIDTH + 'in ' + pc.HEIGHT + 'in;\n' +
    '            margin-top: ' + pc.MARGIN_TOP + 'in;\n' +
    '            margin-bottom: ' + pc.MARGIN_BOTTOM + 'in;\n' +
    '            @bottom-center {\n' +
    '                content: "";\n' +
    '                background-image: none;\n' +
    '            }\n' +
    '        }\n\n' +
    footerCSS + '\n\n' +
    '        @page :right {\n' +
    '            margin-left: ' + pc.MARGIN_INNER + 'in;\n' +
    '            margin-right: ' + pc.MARGIN_OUTER + 'in;\n' +
    '        }\n' +
    '        @page :left {\n' +
    '            margin-left: ' + pc.MARGIN_OUTER + 'in;\n' +
    '            margin-right: ' + pc.MARGIN_INNER + 'in;\n' +
    '        }\n\n' +
    '        /* Base Styles */\n' +
    '        * { margin: 0; padding: 0; box-sizing: border-box; }\n' +
    '        body {\n' +
    '            font-family: \'' + wineFont + '\', \'Georgia\', serif;\n' +
    '            color: ' + col.text + ';\n' +
    '            line-height: 1.4;\n' +
    '        }\n' +
    '        .wine-list { counter-reset: page ' + pc.FRONT_MATTER_PAGES + '; }\n\n' +
    '        /* Title Page */\n' +
    '        .title-page {\n' +
    '            display: flex;\n' +
    '            flex-direction: column;\n' +
    '            justify-content: center;\n' +
    '            align-items: center;\n' +
    '            min-height: 90vh;\n' +
    '            text-align: center;\n' +
    '            page-break-after: always;\n' +
    '        }\n' +
    '        .logo { max-width: 200px; margin-bottom: 40px; }\n' +
    '        .welcome-text {\n' +
    '            max-width: 80%;\n' +
    '            font-family: \'' + wineFont + '\', \'Georgia\', serif;\n' +
    '            font-size: 14px;\n' +
    '            line-height: 1.6;\n' +
    '            color: ' + col.text + ';\n' +
    '        }\n' +
    '        .welcome-text p { margin-bottom: 15px; }\n' +
    '        .welcome-text p:first-child {\n' +
    '            font-family: \'' + (hs[1].title.font.replace(/\.[^.]+$/, '')) + '\', \'Georgia\', serif;\n' +
    '            font-size: 20px;\n' +
    '            color: ' + col.primary + ';\n' +
    '            margin-bottom: 25px;\n' +
    '        }\n' +
    '        .title-date {\n' +
    '            font-size: 10px;\n' +
    '            color: ' + col.text + ';\n' +
    '            margin-top: 40px;\n' +
    '            opacity: 0.6;\n' +
    '        }\n\n' +
    '        /* TOC Page */\n' +
    '        .toc-page {\n' +
    '            page-break-after: always;\n' +
    '            padding-top: 40px;\n' +
    '        }\n' +
    '        .toc-page h2 {\n' +
    '            font-family: \'' + (hs[1].title.font.replace(/\.[^.]+$/, '')) + '\', \'Georgia\', serif;\n' +
    '            font-size: 28px;\n' +
    '            color: ' + col.primary + ';\n' +
    '            text-align: center;\n' +
    '            margin-bottom: 30px;\n' +
    '        }\n' +
    '        .toc-list { list-style: none; columns: 2; column-gap: 40px; }\n' +
    '        .toc-list a {\n' +
    '            text-decoration: none;\n' +
    '            color: ' + col.text + ';\n' +
    '            display: flex;\n' +
    '            justify-content: space-between;\n' +
    '        }\n' +
    '        .toc-page-number { margin-left: 8px; }\n' +
    '        .toc-type-1 { font-weight: bold; font-size: 14px; margin-top: 12px; }\n' +
    '        .toc-type-2 { font-size: 12px; padding-left: 15px; margin-top: 4px; }\n' +
    '        .toc-type-3 { font-size: 11px; padding-left: 30px; margin-top: 2px; }\n\n' +
    '        /* Heading Styles (per type) */' + headingCSS + '\n\n' +
    '        /* Wine Entries */\n' +
    '        .wine-entry {\n' +
    '            display: flex;\n' +
    '            justify-content: space-between;\n' +
    '            align-items: baseline;\n' +
    '            margin: 6px 0;\n' +
    '            padding: 2px 0;\n' +
    '            font-size: ' + we.size + 'px;\n' +
    '            font-family: \'' + wineFont + '\', \'Georgia\', serif;\n' +
    '            font-weight: ' + we.weight + ';\n' +
    '            font-style: ' + we.style + ';\n' +
    '            page-break-inside: avoid;\n' +
    '        }\n' +
    '        .wine-name-vintage {\n' +
    '            flex: 1;\n' +
    '            padding-right: 20px;\n' +
    '        }\n' +
    '        .wine-price {\n' +
    '            text-align: right;\n' +
    '            font-weight: 500;\n' +
    '            white-space: nowrap;\n' +
    '        }\n\n' +
    '        /* Running Labels */' + runningLabelCSS + '\n\n' +
    '        /* Page Breaks */\n' +
    '        .main-content { page: main-pages; }\n' +
    '        .manual-page-break { page-break-before: always; }\n\n' +
    '        @media print {\n' +
    '            body { padding: 0; }\n' +
    '            .title-page, .toc-page { page-break-after: always; }\n' +
    '        }\n' +
    '    </style>\n' +
    '</head>\n' +
    '<body>\n' +
    '    <div class="wine-list">';
}

// ============================================================================
// Title Page
// ============================================================================

function generateTitlePage(logoImageUri, brand) {
  var w = brand.welcome;

  var html = '\n        <div class="title-page">\n' +
    '            <img src="' + logoImageUri + '" alt="Logo" class="logo">\n';

  // Only add welcome text if at least one line has content
  if (w.line1 || w.line2 || w.line3) {
    html += '            <div class="welcome-text">\n';
    if (w.line1) html += '                <p>' + escapeHtml(w.line1) + '</p>\n';
    if (w.line2) html += '                <p>' + escapeHtml(w.line2) + '</p>\n';
    if (w.line3) html += '                <p>' + escapeHtml(w.line3) + '</p>\n';
    html += '            </div>\n';
  }

  // Auto-date
  if (w.showDate) {
    var now = new Date();
    var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var dateStr = Utilities.formatDate(now, timeZone, 'MM.dd.yy');
    html += '            <div class="title-date">' + dateStr + '</div>\n';
  }

  html += '        </div>';
  return html;
}

// ============================================================================
// Table of Contents
// ============================================================================

function generateTOCPage(tocData, brand) {
  var html = '\n        <div class="toc-page">\n' +
    '            <h2>Table of Contents</h2>\n' +
    '            <div class="toc-list">';

  tocData.forEach(function(entry) {
    var pageNum = entry.pageNumber
      ? '<span class="toc-page-number">' + entry.pageNumber + '</span>'
      : '';
    html += '\n                <div class="toc-type-' + entry.type + '">' +
      '<a href="#' + entry.slug + '">' + escapeHtml(entry.title) + pageNum + '</a></div>';
  });

  html += '\n            </div>\n' +
    '        </div>';
  return html;
}

// ============================================================================
// Main Content — The Core Rendering Loop
// ============================================================================

/**
 * Generates the main wine content.
 * Now tracks the current Type 1 section for running labels and injects
 * a running-label div at the top of each page.
 */
function generateMainContent(sections, wineMap, pagination, brand) {
  var showLabel  = brand.footer.showRunningLabel;
  var labelPos   = brand.footer.runningLabelPosition;

  var html = '\n        <main class="main-content">\n' +
    '            <div class="wine-content">';

  // Track current Type 1 section title for running labels
  var currentType1Label = '';
  var pageNumber = 1;  // content page counter for left/right alternation

  for (var s = 0; s < sections.length; s++) {
    var section = sections[s];
    var typeLevel = section.type;
    var hsStyle = brand.headingStyles[typeLevel];

    // Update running label when we hit a Type 1 heading
    if (typeLevel === 1) {
      currentType1Label = section.title;
    }

    // --- Page break before this section? ---
    if (pagination.pageBreaks.has(s)) {
      // Inject footer-position running label before the break (end of current page)
      if (showLabel && labelPos === 'footer' && currentType1Label) {
        var footerAlign = (pageNumber % 2 === 0) ? 'left' : 'right';
        html += '\n                <div class="running-label running-label-' + footerAlign + '">' +
          escapeHtml(currentType1Label) + '</div>';
      }

      html += '\n                <div class="manual-page-break"></div>';
      pageNumber++;

      // Inject header-position running label after the break (top of new page)
      if (showLabel && labelPos === 'header' && currentType1Label) {
        var headerAlign = (pageNumber % 2 === 0) ? 'left' : 'right';
        html += '\n                <div class="running-label running-label-' + headerAlign + '">' +
          escapeHtml(currentType1Label) + '</div>';
      }
    }

    // --- Heading ---
    var slug = createSlug('section-' + typeLevel + '-' + section.title);

    if (hsStyle && hsStyle.subtext.position === 'inline' && section.subtext) {
      html += '\n                <div class="heading-type-' + typeLevel + '" id="' + slug + '">' +
        escapeHtml(section.title) +
        ' <span class="subtext-type-' + typeLevel + '">' + escapeHtml(section.subtext) + '</span>' +
        '</div>';
    } else {
      html += '\n                <div class="heading-type-' + typeLevel + '" id="' + slug + '">' +
        escapeHtml(section.title) + '</div>';

      if (section.subtext) {
        html += '\n                <div class="subtext-type-' + typeLevel + '">' +
          escapeHtml(section.subtext) + '</div>';
      }
    }

    // --- Wines for this section ---
    var wines = (section.code > 0 && wineMap.has(section.code)) ? wineMap.get(section.code) : [];
    var wineBreaks = pagination.winePageBreaks.get(s);

    for (var w = 0; w < wines.length; w++) {
      if (wineBreaks && wineBreaks.has(w)) {
        // Running label before wine-level page break
        if (showLabel && labelPos === 'footer' && currentType1Label) {
          var wfAlign = (pageNumber % 2 === 0) ? 'left' : 'right';
          html += '\n                <div class="running-label running-label-' + wfAlign + '">' +
            escapeHtml(currentType1Label) + '</div>';
        }

        html += '\n                <div class="manual-page-break"></div>';
        pageNumber++;

        if (showLabel && labelPos === 'header' && currentType1Label) {
          var whAlign = (pageNumber % 2 === 0) ? 'left' : 'right';
          html += '\n                <div class="running-label running-label-' + whAlign + '">' +
            escapeHtml(currentType1Label) + '</div>';
        }
      }

      var wine = wines[w];
      var vintage = wine.vintage ? ' ' + wine.vintage : '';
      var price = wine.price ? Math.round(wine.price).toString() : '';

      html += '\n                <div class="wine-entry">' +
        '<span class="wine-name-vintage">' + escapeHtml(wine.name) + escapeHtml(vintage) + '</span>' +
        '<span class="wine-price">' + price + '</span>' +
        '</div>';
    }
  }

  html += '\n            </div>\n        </main>';
  return html;
}

// ============================================================================
// HTML Foot
// ============================================================================

function generateHTMLFoot() {
  return '\n    </div>\n</body>\n</html>';
}

/**
 * Creates and displays a download dialog for the generated HTML.
 */
function showDownloadDialog(html) {
  var encodedHtml = Utilities.base64Encode(html, Utilities.Charset.UTF_8);
  var dataUri = 'data:text/html;charset=utf-8;base64,' + encodedHtml;

  var dialogHtml = '<div style="font-family: Arial, sans-serif; padding: 20px;">' +
    '<h2>Your Wine List is Ready!</h2>' +
    '<a href="' + dataUri + '" download="wine_list.html" ' +
    'style="display: inline-block; margin-top: 15px; padding: 10px 20px; ' +
    'background: #4285f4; color: white; text-decoration: none; border-radius: 4px;">' +
    'Download Wine List</a>' +
    '<p style="margin-top: 20px; font-size: 12px; color: #999;">' +
    'Note: Open the HTML file in a browser and use Print → Save as PDF for best results.</p>' +
    '</div>';

  var output = HtmlService.createHtmlOutput(dialogHtml)
    .setWidth(450)
    .setHeight(250);

  SpreadsheetApp.getUi().showModalDialog(output, 'Download Wine List');
}
