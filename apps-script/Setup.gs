/**
 * One-time (and re-runnable) workbook setup: creates every tab with the
 * exact columns from the spec, header styling, number formats, dropdown
 * validation, and the auto Financial Year / Month formulas.
 *
 * Safe to re-run: it only (re)writes headers/formats/validation, it never
 * clears existing data rows.
 */

function setupWorkbook() {
  var ss = getSs();

  setupTransactionsSheet(ss);
  setupBudgetSheet(ss);
  setupSavingsGoalsSheet(ss);
  setupSavingsContributionsSheet(ss);
  setupReconciliationSheet(ss);
  setupDashboardSheet(ss);

  orderSheets(ss);
  removeDefaultBlankSheet(ss);

  SpreadsheetApp.flush();
  ss.toast('Budget Tracker tabs are set up.', 'Setup complete', 5);
}

function setupTransactionsSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.TRANSACTIONS);
  var headers = ['Date', 'Amount', 'Direction', 'Category', 'Scope', 'Note',
    'Receipt Link', 'Financial Year', 'Month', 'Reconciled'];
  setHeaderRow(sheet, headers);

  sheet.getRange('A2:A').setNumberFormat(DATE_FORMAT);
  sheet.getRange('B2:B').setNumberFormat(CURRENCY_FORMAT);
  sheet.getRange('I2:I').setNumberFormat(MONTH_FORMAT);

  applyListValidation(sheet.getRange('C2:C'), DIRECTIONS, true);
  applyListValidation(sheet.getRange('D2:D'), ALL_CATEGORIES, false); // blank allowed -> flagged Uncategorised (Layer 2)
  applyListValidation(sheet.getRange('E2:E'), SCOPES, true);
  applyListValidation(sheet.getRange('J2:J'), RECONCILED_OPTIONS, true);

  // Auto-fill Financial Year (H) and Month (I) from Date (A). Single
  // ARRAYFORMULA anchored at row 2 so it keeps up with every new row,
  // whether added via the sidebar form or typed directly.
  sheet.getRange('H2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",' +
    '"FY"&IF(MONTH(A2:A)>=7,YEAR(A2:A),YEAR(A2:A)-1)' +
    '&"-"&RIGHT(IF(MONTH(A2:A)>=7,YEAR(A2:A)+1,YEAR(A2:A)),2)))'
  );
  sheet.getRange('I2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",DATE(YEAR(A2:A),MONTH(A2:A),1)))'
  );

  protectFormulaColumn(sheet, 8, 1, 'Auto-filled: Financial Year');
  protectFormulaColumn(sheet, 9, 1, 'Auto-filled: Month');

  sheet.setColumnWidths(1, 10, 130);
}

function setupBudgetSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.BUDGET);
  setHeaderRow(sheet, ['Month', 'Category', 'Budgeted Amount']);

  sheet.getRange('A2:A').setNumberFormat(MONTH_FORMAT);
  sheet.getRange('C2:C').setNumberFormat(CURRENCY_FORMAT);
  applyListValidation(sheet.getRange('B2:B'), EXPENSE_CATEGORIES, true);

  sheet.setColumnWidths(1, 3, 160);
}

function setupSavingsGoalsSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.SAVINGS_GOALS);
  setHeaderRow(sheet, ['Goal Name', 'Target', 'Deadline', 'Current']);

  sheet.getRange('B2:B').setNumberFormat(CURRENCY_FORMAT);
  sheet.getRange('C2:C').setNumberFormat(DATE_FORMAT);
  sheet.getRange('D2:D').setNumberFormat(CURRENCY_FORMAT);

  sheet.setColumnWidths(1, 4, 160);
}

function setupSavingsContributionsSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.SAVINGS_CONTRIBUTIONS);
  setHeaderRow(sheet, ['Date', 'Goal Name', 'Amount']);

  sheet.getRange('A2:A').setNumberFormat(DATE_FORMAT);
  sheet.getRange('C2:C').setNumberFormat(CURRENCY_FORMAT);

  var goalsSheet = ss.getSheetByName(SHEET_NAMES.SAVINGS_GOALS);
  if (goalsSheet) {
    var goalsRange = goalsSheet.getRange('A2:A1000');
    applyRangeValidation(sheet.getRange('B2:B'), goalsRange, false);
  }

  sheet.setColumnWidths(1, 3, 160);
}

function setupReconciliationSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.RECONCILIATION);
  setHeaderRow(sheet, ['Month', 'Actual Bank Net', 'Tracked Net', 'Difference']);

  sheet.getRange('A2:A').setNumberFormat(MONTH_FORMAT);
  sheet.getRange('B2:D').setNumberFormat(CURRENCY_FORMAT);

  sheet.setColumnWidths(1, 4, 160);
}

function setupDashboardSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.DASHBOARD);
  sheet.clear();
  sheet.getRange('A1').setValue('Budget Tracker Dashboard')
    .setFontWeight('bold').setFontSize(16);
  sheet.getRange('A2').setValue('Formulas and charts are added in later build layers.')
    .setFontStyle('italic').setFontColor('#666666');
  sheet.setColumnWidth(1, 220);
}

function protectFormulaColumn(sheet, column, numColumns, description) {
  var range = sheet.getRange(2, column, sheet.getMaxRows() - 1, numColumns);
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === description) {
      protections[i].remove();
    }
  }
  var protection = range.protect().setDescription(description);
  protection.setWarningOnly(true); // warns editors instead of hard-blocking
}

function orderSheets(ss) {
  for (var i = 0; i < SHEET_ORDER.length; i++) {
    var sheet = ss.getSheetByName(SHEET_ORDER[i]);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }
}

function removeDefaultBlankSheet(ss) {
  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    var range = defaultSheet.getDataRange();
    var isEmpty = range.getNumRows() === 1 && range.getNumColumns() === 1 && !range.getValue();
    if (isEmpty) {
      ss.deleteSheet(defaultSheet);
    }
  }
}
