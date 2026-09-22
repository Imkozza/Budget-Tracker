/**
 * Shared helpers used across the Budget Tracker project.
 */

function getSs() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  var sheet = getSs().getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" not found. Run "Budget Tracker > Run Setup" from the menu first.');
  }
  return sheet;
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function setHeaderRow(sheet, headers) {
  var range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
  sheet.setFrozenRows(1);
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Applies a dropdown data-validation rule to a range.
 * @param {Range} range Target range (usually a full column from row 2 down).
 * @param {Array<string>} items Allowed values.
 * @param {boolean} strict If true, invalid/blank entries are rejected; if false, a warning is shown but entry is allowed.
 */
function applyListValidation(range, items, strict) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(!strict)
    .build();
  range.setDataValidation(rule);
}

function applyRangeValidation(range, sourceRange, strict) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(!strict)
    .build();
  range.setDataValidation(rule);
}

/**
 * Finds the next blank row based on a specific "key" column, rather than
 * relying on getLastRow(), because columns filled by an ARRAYFORMULA (e.g.
 * auto Financial Year / Month) make getLastRow() see the whole column as
 * "used" even where the key column is blank.
 */
function getNextDataRow(sheet, keyColumn) {
  var numRows = sheet.getMaxRows() - 1;
  if (numRows <= 0) return 2;
  var values = sheet.getRange(2, keyColumn, numRows, 1).getValues();
  var last = 1;
  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v !== '' && v !== null) {
      last = i + 2;
    }
  }
  return last + 1;
}

function parseFormDate(dateStr) {
  // dateStr from an <input type="date"> is "yyyy-mm-dd" in the user's locale-less form.
  var parts = dateStr.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

/** AU financial year label for a date, e.g. 1 Aug 2025 -> "FY2025-26". */
function financialYearLabel(date) {
  var year = date.getFullYear();
  var month = date.getMonth() + 1; // 1-12
  var startYear = month >= 7 ? year : year - 1;
  var endYearShort = String(startYear + 1).slice(-2);
  return 'FY' + startYear + '-' + endYearShort;
}

/** First-of-month Date for grouping, e.g. 2025-08-17 -> 2025-08-01. */
function monthStart(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}
