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

function generateHTML(sections, wineMap, pagination, assets, brand) {
  var parts = [generateHTMLHead(assets, brand)];

  if (brand.welcome.showTitlePage) {
    parts.push(generateTitlePage(assets.logoImageUri, brand));
  }

  parts.push(generateTOCPage(pagination.tocData, brand));
  parts.push(generateMainContent(sections, wineMap, pagination, brand));
  parts.push(generateHTMLFoot());
  return parts.join('');
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
  var wineFont = we.font.replace(/\.[^.]+$/, '');

  // --- @font-face declarations ---
  var fontFaceParts = [];
  assets.fontUris.forEach(function(fontData, fileName) {
    var familyName = fileName.replace(/\.[^.]+$/, '');
    fontFaceParts.push(
      '\n        @font-face {\n' +
      '            font-family: \'' + familyName + '\';\n' +
      '            src: url(\'' + fontData.uri + '\') format(\'' + fontData.format + '\');\n' +
      '            font-weight: normal;\n' +
      '            font-style: normal;\n' +
      '            font-display: block;\n' +
      '        }'
    );
  });

  // --- Per-type heading CSS ---
  var headingCSSParts = [];
  for (var t = 1; t <= MAX_HEADING_TYPE; t++) {
    var ts = hs[t].title;
    var ss = hs[t].subtext;
    var titleFont = ts.font.replace(/\.[^.]+$/, '');
    var subFont   = ss.font.replace(/\.[^.]+$/, '');

    var fontVariant = (ts.variant && ts.variant !== 'normal' && ts.transform === 'none')
      ? ts.variant
      : 'normal';

    var marginTop    = (ts.spaceBefore !== null && ts.spaceBefore !== undefined)
      ? ts.spaceBefore : Math.round(ts.size * 0.7);
    var marginBottom = (ts.spaceAfter  !== null && ts.spaceAfter  !== undefined)
      ? ts.spaceAfter  : Math.round(ts.size * 0.4);

    headingCSSParts.push(
      '\n        .heading-type-' + t + ' {\n' +
      '            font-family: \'' + titleFont + '\', \'Georgia\', serif;\n' +
      '            font-size: ' + ts.size + 'px;\n' +
      '            color: ' + ts.color + ';\n' +
      '            text-align: ' + ts.align + ';\n' +
      '            font-weight: ' + ts.weight + ';\n' +
      '            text-transform: ' + ts.transform + ';\n' +
      '            font-variant: ' + fontVariant + ';\n' +
      '            font-variant-caps: ' + fontVariant + ';\n' +
      '            letter-spacing: ' + ts.spacing + 'px;\n' +
      '            margin: ' + marginTop + 'px 0 ' + marginBottom + 'px 0;\n' +
      '            page-break-after: avoid;\n' +
      '            scroll-margin-top: 20px;\n' +
      '        }'
    );

    if (ts.underline === 'partial') {
      headingCSSParts.push(
        '\n        .heading-type-' + t + '::after {\n' +
        '            content: \'\';\n' +
        '            position: absolute;\n' +
        '            bottom: -2px;\n' +
        '            left: 0;\n' +
        '            width: 85%;\n' +
        '            height: 1px;\n' +
        '            background-color: ' + ts.color + ';\n' +
        '        }\n' +
        '        .heading-type-' + t + ' { position: relative; }'
      );
    } else if (ts.underline === 'text') {
      headingCSSParts.push(
        '\n        .heading-type-' + t + ' {\n' +
        '            display: inline-block;\n' +
        '            border-bottom: 1px solid ' + ts.color + ';\n' +
        '        }'
      );
    } else if (ts.underline === 'full') {
      headingCSSParts.push(
        '\n        .heading-type-' + t + '::after {\n' +
        '            content: \'\';\n' +
        '            position: absolute;\n' +
        '            bottom: -2px;\n' +
        '            left: 0;\n' +
        '            width: 100%;\n' +
        '            height: 1px;\n' +
        '            background-color: ' + ts.color + ';\n' +
        '        }\n' +
        '        .heading-type-' + t + ' { position: relative; }'
      );
    }

    var subtextBlock = '\n        .subtext-type-' + t + ' {\n' +
      '            font-family: \'' + subFont + '\', \'Georgia\', serif;\n' +
      '            font-size: ' + ss.size + 'px;\n' +
      '            color: ' + ss.color + ';\n' +
      '            font-weight: ' + ss.weight + ';\n' +
      '            text-align: ' + ts.align + ';\n';
    subtextBlock += (ss.position === 'inline')
      ? '            display: inline;\n            margin-left: 8px;\n'
      : '            display: block;\n            margin-top: 2px;\n            margin-bottom: 6px;\n';
    subtextBlock += '        }';
    headingCSSParts.push(subtextBlock);
  }

// --- Footer CSS ---
  var hasImage = !!(assets.footerImageUri && assets.footerImageUri.length > 0);

  var numStyle =
    'font-family: \'' + wineFont + '\', \'Georgia\', serif; ' +
    'font-size: 10px; ' +
    'color: ' + col.primary + '; ' +
    'vertical-align: middle;';

  var ruleBorder = '';
  var rulePadding = 'padding-top: 6px;';
  if (foot.footerRule === 'single') {
    ruleBorder = 'border-top: 1px solid ' + col.primary + ';';
  } else if (foot.footerRule === 'double') {
    ruleBorder = 'border-top: 3px double ' + col.primary + ';';
  }
  var ruleStyle     = ruleBorder + ' ' + (foot.footerRule !== 'none' ? rulePadding : '');
  var ruleStyleFull = ruleStyle.trim();

  var imageBg = hasImage
    ? 'background-image: url(\'' + assets.footerImageUri + '\'); ' +
      'background-repeat: no-repeat; ' +
      'background-position: center center; ' +
      'background-size: 90% auto;'
    : '';

  var footerCSS;

  if (foot.pageNumberPosition === 'outer') {
    footerCSS =
      '\n        @page main-pages {\n' +
      '            @bottom-left   { content: ""; ' + ruleStyleFull + ' }\n' +
      '            @bottom-center {\n' +
      '                content: "";\n' +
      '                ' + imageBg + '\n' +
      '                ' + ruleStyleFull + '\n' +
      '                height: 0.4in;\n' +
      '                width: 100%;\n' +
      '            }\n' +
      '            @bottom-right  { content: ""; ' + ruleStyleFull + ' }\n' +
      '        }\n' +
      '        @page main-pages:right {\n' +
      '            @bottom-left  { content: ""; ' + ruleStyleFull + ' }\n' +
      '            @bottom-right {\n' +
      '                content: counter(page);\n' +
      '                ' + numStyle + '\n' +
      '                ' + ruleStyleFull + '\n' +
      '                height: 0.4in;\n' +
      '                padding-top: 0.15in;\n' +
      '                text-align: right;\n' +
      '            }\n' +
      '        }\n' +
      '        @page main-pages:left {\n' +
      '            @bottom-left {\n' +
      '                content: counter(page);\n' +
      '                ' + numStyle + '\n' +
      '                ' + ruleStyleFull + '\n' +
      '                height: 0.4in;\n' +
      '                padding-top: 0.15in;\n' +
      '                text-align: left;\n' +
      '            }\n' +
      '            @bottom-right { content: ""; ' + ruleStyleFull + ' }\n' +
      '        }';

  } else {
    // Center: background-image for decorative image, content for page number — both in @bottom-center
    footerCSS =
      '\n        @page main-pages {\n' +
      '            @bottom-left  { content: ""; ' + ruleStyleFull + ' }\n' +
      '            @bottom-center {\n' +
      '                content: counter(page);\n' +
      '                ' + numStyle + '\n' +
      '                ' + imageBg + '\n' +
      '                background-origin: content-box;\n' +
      '                ' + ruleStyleFull + '\n' +
      '                height: 0.4in;\n' +
      '                padding-top: 0.15in;\n' +
      '                padding-bottom: 0.1in;\n' +
      '                width: 100%;\n' +
      '                text-align: center;\n' +
      '            }\n' +
      '            @bottom-right { content: ""; ' + ruleStyleFull + ' }\n' +
      '        }';
  }

  // --- Running label CSS ---
  var runningLabelCSS = '';
  if (foot.showRunningLabel) {
    if (foot.runningLabelPosition === 'header') {
      runningLabelCSS =
        '\n        .running-label {\n' +
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
      runningLabelCSS =
        '\n        .running-label {\n' +
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

  // --- Assemble full output ---
  var titleHeadingFont = hs[1].title.font.replace(/\.[^.]+$/, '');
  var out = [
    '<!DOCTYPE html>\n<html lang="en">\n<head>\n',
    '    <meta charset="UTF-8">\n',
    '    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n',
    '    <title>Wine List</title>\n',
    '    <style>\n',
    '        /* Embedded Fonts */', fontFaceParts.join(''), '\n\n',
    '        /* Page Setup */\n',
    '        @page {\n',
    '            size: ', pc.WIDTH, 'in ', pc.HEIGHT, 'in;\n',
    '            margin-top: ', pc.MARGIN_TOP, 'in;\n',
    '            margin-bottom: ', pc.MARGIN_BOTTOM, 'in;\n',
    '            @bottom-center { content: ""; background-image: none; }\n',
    '        }\n\n',
    footerCSS, '\n\n',
    '        @page :right {\n',
    '            margin-left: ', pc.MARGIN_INNER, 'in;\n',
    '            margin-right: ', pc.MARGIN_OUTER, 'in;\n',
    '        }\n',
    '        @page :left {\n',
    '            margin-left: ', pc.MARGIN_OUTER, 'in;\n',
    '            margin-right: ', pc.MARGIN_INNER, 'in;\n',
    '        }\n\n',
    '        /* Base Styles */\n',
    '        * { margin: 0; padding: 0; box-sizing: border-box; }\n',
    '        body {\n',
    '            font-family: \'', wineFont, '\', \'Georgia\', serif;\n',
    '            color: ', col.text, ';\n',
    '            line-height: ', ELEMENT_HEIGHTS.LINE_HEIGHT, ';\n',
    '        }\n',
    '        .wine-list { counter-reset: page ', pc.FRONT_MATTER_PAGES, '; }\n\n',
    '        /* Title Page */\n',
    '        .title-page {\n',
    '            display: flex;\n',
    '            flex-direction: column;\n',
    '            justify-content: center;\n',
    '            align-items: center;\n',
    '            min-height: 90vh;\n',
    '            text-align: center;\n',
    '            page-break-after: always;\n',
    '        }\n',
    '        .logo { max-width: 200px; margin-bottom: 40px; }\n',
    '        .welcome-text {\n',
    '            max-width: 80%;\n',
    '            font-family: \'', wineFont, '\', \'Georgia\', serif;\n',
    '            font-size: 14px;\n',
    '            line-height: 1.6;\n',
    '            color: ', col.text, ';\n',
    '        }\n',
    '        .welcome-text p { margin-bottom: 15px; }\n',
    '        .welcome-text p:first-child {\n',
    '            font-family: \'', titleHeadingFont, '\', \'Georgia\', serif;\n',
    '            font-size: 20px;\n',
    '            font-weight: ', hs[1].title.weight, ';\n',
    '            color: ', col.primary, ';\n',
    '            margin-bottom: 25px;\n',
    '        }\n',
    '        .title-date {\n',
    '            font-size: 10px;\n',
    '            color: ', col.text, ';\n',
    '            margin-top: 40px;\n',
    '            opacity: 0.6;\n',
    '        }\n\n',
    '        /* TOC Page */\n',
    '        .toc-page { page-break-after: always; padding-top: 40px; }\n',
    '        .toc-page h2 {\n',
    '            font-family: \'', titleHeadingFont, '\', \'Georgia\', serif;\n',
    '            font-size: 28px;\n',
    '            font-weight: ', hs[1].title.weight, ';\n',
    '            color: ', col.primary, ';\n',
    '            text-align: center;\n',
    '            margin-bottom: 30px;\n',
    '        }\n',
    '        .toc-list { list-style: none; columns: 2; column-gap: 40px; }\n',
    '        .toc-list a {\n',
    '            text-decoration: none;\n',
    '            color: ', col.text, ';\n',
    '            display: flex;\n',
    '            justify-content: space-between;\n',
    '        }\n',
    '        .toc-page-number { margin-left: 8px; }\n',
    '        .toc-type-1 {\n' +
    '            font-size: 14px;\n' +
    '            margin-top: 12px;\n' +
    '            font-family: \'' + hs[2].title.font.replace(/\.[^.]+$/, '') + '\', \'Georgia\', serif;\n' +
    '            font-weight: ' + hs[2].title.weight + ';\n' +
    '            text-transform: ' + hs[2].title.transform + ';\n' +
    '            font-variant-caps: ' + (hs[2].title.variant !== 'normal' && hs[2].title.transform === 'none' ? hs[2].title.variant : 'normal') + ';\n' +
    '            letter-spacing: ' + hs[2].title.spacing + 'px;\n' +
    '        }\n' +
    '        .toc-type-2 {\n' +
    '            font-size: 12px;\n' +
    '            padding-left: 15px;\n' +
    '            margin-top: 4px;\n' +
    '            font-family: \'' + hs[2].title.font.replace(/\.[^.]+$/, '') + '\', \'Georgia\', serif;\n' +
    '            font-weight: ' + hs[2].title.weight + ';\n' +
    '            text-transform: ' + hs[2].title.transform + ';\n' +
    '            font-variant-caps: ' + (hs[2].title.variant !== 'normal' && hs[2].title.transform === 'none' ? hs[2].title.variant : 'normal') + ';\n' +
    '            letter-spacing: ' + hs[2].title.spacing + 'px;\n' +
    '        }\n' +
    '        .toc-type-3 {\n' +
    '            font-size: 11px;\n' +
    '            padding-left: 30px;\n' +
    '            margin-top: 2px;\n' +
    '            font-family: \'' + hs[2].title.font.replace(/\.[^.]+$/, '') + '\', \'Georgia\', serif;\n' +
    '            font-weight: ' + hs[2].title.weight + ';\n' +
    '            text-transform: ' + hs[2].title.transform + ';\n' +
    '            font-variant-caps: ' + (hs[2].title.variant !== 'normal' && hs[2].title.transform === 'none' ? hs[2].title.variant : 'normal') + ';\n' +
    '            letter-spacing: ' + hs[2].title.spacing + 'px;\n' +
    '        }\n\n',
    '        /* Heading Styles (per type) */', headingCSSParts.join(''), '\n\n',
    '        /* Wine Entries */\n',
    '        .wine-entry {\n',
    '            display: flex;\n',
    '            justify-content: space-between;\n',
    '            align-items: baseline;\n',
    '            margin: ', ELEMENT_HEIGHTS.WINE_ENTRY_MARGIN, 'px 0;\n',
    '            padding: ', ELEMENT_HEIGHTS.WINE_ENTRY_PADDING, 'px 0;\n',
    '            font-size: ', we.size, 'px;\n',
    '            font-family: \'', wineFont, '\', \'Georgia\', serif;\n',
    '            font-weight: ', we.weight, ';\n',
    '            font-style: ', we.style, ';\n',
    '            page-break-inside: avoid;\n',
    '        }\n',
    '        .wine-name-vintage { flex: 1; padding-right: 20px; }\n',
    '        .wine-price { text-align: right; font-weight: 500; white-space: nowrap; }\n\n',
    '        /* Running Labels */', runningLabelCSS, '\n\n',
    '        /* Page Breaks */\n',
    '        .main-content { page: main-pages; }\n',
    '        .manual-page-break { page-break-before: always; }\n\n',
    '        @media print {\n',
    '            body { padding: 0; }\n',
    '            .title-page, .toc-page { page-break-after: always; }\n',
    '        }\n',
    '    </style>\n',
    '</head>\n',
    '<body>\n',
    '    <div class="wine-list">'
  ];

  return out.join('');
}

// ============================================================================
// Title Page
// ============================================================================

function generateTitlePage(logoImageUri, brand) {
  var w    = brand.welcome;
  var out  = ['\n        <div class="title-page">\n'];

  out.push('            <img src="' + logoImageUri + '" alt="Logo" class="logo">\n');

  if (w.line1 || w.line2 || w.line3) {
    out.push('            <div class="welcome-text">\n');
    if (w.line1) out.push('                <p>', escapeHtml(w.line1), '</p>\n');
    if (w.line2) out.push('                <p>', escapeHtml(w.line2), '</p>\n');
    if (w.line3) out.push('                <p>', escapeHtml(w.line3), '</p>\n');
    out.push('            </div>\n');
  }

  if (w.showDate) {
    var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var dateStr  = Utilities.formatDate(new Date(), timeZone, 'MM.dd.yy');
    out.push('            <div class="title-date">', dateStr, '</div>\n');
  }

  out.push('        </div>');
  return out.join('');
}

// ============================================================================
// Table of Contents
// ============================================================================

function generateTOCPage(tocData, brand) {
  var out = [
    '\n        <div class="toc-page">\n',
    '            <h2>Table of Contents</h2>\n',
    '            <div class="toc-list">'
  ];

  tocData.forEach(function(entry) {
    var pageNum = entry.pageNumber
      ? '<span class="toc-page-number">' + entry.pageNumber + '</span>'
      : '';
    out.push(
      '\n                <div class="toc-type-', entry.type, '">',
      '<a href="#', entry.slug, '">', escapeHtml(entry.title), pageNum, '</a></div>'
    );
  });

  out.push('\n            </div>\n        </div>');
  return out.join('');
}

// ============================================================================
// Main Content — The Core Rendering Loop
// ============================================================================

function generateMainContent(sections, wineMap, pagination, brand) {
  var showLabel = brand.footer.showRunningLabel;
  var labelPos  = brand.footer.runningLabelPosition;

  var out = [
    '\n        <main class="main-content">\n',
    '            <div class="wine-content">'
  ];

  var currentType1Label = '';
  var pageNumber        = 1;   // content-page counter for left/right label alternation

  for (var s = 0; s < sections.length; s++) {
    var section   = sections[s];
    var typeLevel = section.type;
    var hsStyle   = brand.headingStyles[typeLevel];

    if (typeLevel === 1) currentType1Label = section.title;

    // --- Section-level page break ---
    if (pagination.pageBreaks.has(s)) {
      if (showLabel && labelPos === 'footer' && currentType1Label) {
        var fa = (pageNumber % 2 === 0) ? 'left' : 'right';
        out.push('\n                <div class="running-label running-label-', fa, '">',
          escapeHtml(currentType1Label), '</div>');
      }
      out.push('\n                <div class="manual-page-break"></div>');
      pageNumber++;
      if (showLabel && labelPos === 'header' && currentType1Label) {
        var ha = (pageNumber % 2 === 0) ? 'left' : 'right';
        out.push('\n                <div class="running-label running-label-', ha, '">',
          escapeHtml(currentType1Label), '</div>');
      }
    }

    // --- Heading ---
    var slug = createSlug('section-' + typeLevel + '-' + section.title);
    if (hsStyle && hsStyle.subtext.position === 'inline' && section.subtext) {
      out.push(
        '\n                <div class="heading-type-', typeLevel, '" id="', slug, '">',
        escapeHtml(section.title),
        ' <span class="subtext-type-', typeLevel, '">', escapeHtml(section.subtext), '</span>',
        '</div>'
      );
    } else {
      out.push(
        '\n                <div class="heading-type-', typeLevel, '" id="', slug, '">',
        escapeHtml(section.title), '</div>'
      );
      if (section.subtext) {
        out.push(
          '\n                <div class="subtext-type-', typeLevel, '">',
          escapeHtml(section.subtext), '</div>'
        );
      }
    }

    // --- Wine entries ---
    var wines      = (section.code > 0 && wineMap.has(section.code)) ? wineMap.get(section.code) : [];
    var wineBreaks = pagination.winePageBreaks.get(s);

    for (var w = 0; w < wines.length; w++) {

      // Wine-level page break
      if (wineBreaks && wineBreaks.has(w)) {
        if (showLabel && labelPos === 'footer' && currentType1Label) {
          var wfa = (pageNumber % 2 === 0) ? 'left' : 'right';
          out.push('\n                <div class="running-label running-label-', wfa, '">',
            escapeHtml(currentType1Label), '</div>');
        }
        out.push('\n                <div class="manual-page-break"></div>');
        pageNumber++;
        if (showLabel && labelPos === 'header' && currentType1Label) {
          var wha = (pageNumber % 2 === 0) ? 'left' : 'right';
          out.push('\n                <div class="running-label running-label-', wha, '">',
            escapeHtml(currentType1Label), '</div>');
        }
      }

      var wine    = wines[w];
      var vintage = wine.vintage ? ' ' + wine.vintage : '';
      var price   = wine.price   ? Math.round(wine.price).toString() : '';

      out.push(
        '\n                <div class="wine-entry">',
        '<span class="wine-name-vintage">', escapeHtml(wine.name), escapeHtml(vintage), '</span>',
        '<span class="wine-price">', price, '</span>',
        '</div>'
      );
    }
  }

  out.push('\n            </div>\n        </main>');
  return out.join('');
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
