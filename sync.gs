// sync.gs

// ============================================================================
// Section Code Sync
// ============================================================================

/**
 * Scans Sections!A:C to build a map of section title → current full code
 * (e.g. "Champagne" → "003-Champagne"), then updates every row in List!B
 * whose name-part matches but whose prefix is stale.
 *
 * Sections!A contains formula-generated values in the format "001-Champagne".
 * Sections!C contains the raw title ("Champagne") used for matching.
 * List!B  contains the section code assigned to each wine.
 *
 * @returns {Object} { success: boolean, message: string }
 */
function syncSectionCodes() {
  var ss             = SpreadsheetApp.getActiveSpreadsheet();
  var sectionsSheet  = ss.getSheetByName(SHEETS.SECTIONS);
  var listSheet      = ss.getSheetByName(SHEETS.LIST);

  if (!sectionsSheet || !listSheet) {
    SpreadsheetApp.getUi().alert('Error: "Sections" or "List" sheet not found.');
    return { success: false, message: 'Sheet not found.' };
  }

  // ── 1. Build title → currentCode map from Sections A & C ──────────────────
  var sectionsLastRow = sectionsSheet.getLastRow();
  if (sectionsLastRow < 2) {
    SpreadsheetApp.getUi().alert('No data found in the Sections sheet.');
    return { success: false, message: 'Sections sheet is empty.' };
  }

  // Columns A (index 0) = full code, C (index 2) = title
  var sectionsData = sectionsSheet
    .getRange(2, 1, sectionsLastRow - 1, 3)
    .getValues();

  // Map: normalised title → current full code string
  var titleToCode = {};
  sectionsData.forEach(function(row) {
    var fullCode = row[0].toString().trim();   // e.g. "003-Champagne"
    var title    = row[2].toString().trim();   // e.g. "Champagne"
    if (title && fullCode) {
      titleToCode[title.toLowerCase()] = fullCode;
    }
  });

  if (Object.keys(titleToCode).length === 0) {
    SpreadsheetApp.getUi().alert('No valid section codes found in the Sections sheet.');
    return { success: false, message: 'No section codes found.' };
  }

  // ── 2. Locate the Section column in List ───────────────────────────────────
  var listLastRow = listSheet.getLastRow();
  var listLastCol = listSheet.getLastColumn();
  if (listLastRow < 2 || listLastCol < 1) {
    SpreadsheetApp.getUi().alert('No data found in the List sheet.');
    return { success: false, message: 'List sheet is empty.' };
  }

  var listHeaders    = listSheet.getRange(1, 1, 1, listLastCol).getValues()[0];
  var sectionColIdx  = listHeaders.indexOf('Section'); // 0-based
  if (sectionColIdx < 0) {
    SpreadsheetApp.getUi().alert('Column "Section" not found in the List sheet.');
    return { success: false, message: '"Section" column not found.' };
  }

  // ── 3. Read all Section values in one call ─────────────────────────────────
  var dataRowCount  = listLastRow - 1;
  var sectionRange  = listSheet.getRange(2, sectionColIdx + 1, dataRowCount, 1);
  var sectionValues = sectionRange.getValues(); // [ [val], [val], ... ]

  // ── 4. Compute corrections ─────────────────────────────────────────────────
  var updatedCount = 0;
  var unmatched    = [];

  var newValues = sectionValues.map(function(row, i) {
    var cellValue = row[0].toString().trim();
    if (!cellValue) return row; // blank — leave alone

    // Parse the name portion after the first dash (handles "003-Champagne")
    var dashIdx  = cellValue.indexOf('-');
    var namePart = (dashIdx >= 0)
      ? cellValue.substring(dashIdx + 1).trim()
      : cellValue;

    var correctCode = titleToCode[namePart.toLowerCase()];

    if (!correctCode) {
      // Section title not found in Sections sheet — flag it, leave value alone
      unmatched.push('"' + cellValue + '" (row ' + (i + 2) + ')');
      return row;
    }

    if (correctCode !== cellValue) {
      updatedCount++;
      return [correctCode];
    }

    return row; // already correct
  });

  // ── 5. Write back in a single batch call ──────────────────────────────────
  if (updatedCount > 0) {
    sectionRange.setValues(newValues);
  }

  // ── 6. Report results ──────────────────────────────────────────────────────
  var msg = '';

  if (updatedCount === 0 && unmatched.length === 0) {
    msg = 'All section codes are already up to date. No changes needed.';
  } else {
    if (updatedCount > 0) {
      msg += updatedCount + ' section code' + (updatedCount === 1 ? '' : 's') + ' updated.\n';
    } else {
      msg += 'No section codes needed updating.\n';
    }
    if (unmatched.length > 0) {
      msg += '\nThe following ' + unmatched.length +
             ' entr' + (unmatched.length === 1 ? 'y' : 'ies') +
             ' could not be matched to any current section title and were left unchanged:\n• ' +
             unmatched.join('\n• ');
    }
  }

  SpreadsheetApp.getUi().alert(msg);
  return { success: true, message: msg };
}
