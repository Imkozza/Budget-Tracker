/**
 * One-time (and re-runnable) workbook setup: creates every tab with the
 * exact columns from the spec, header styling, number formats, dropdown
 * validation, auto formulas, and the accountability-rule conditional
 * formatting / flags.
 *
 * Safe to re-run: it only (re)writes headers/formats/validation/formulas,
 * it never clears existing data rows. Conditional format rules are fully
 * replaced (not appended) each run, so re-running never duplicates them.
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
  getOrCreateReceiptsFolder(); // ensure the one Drive folder exists up front

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
  applyListValidation(sheet.getRange('D2:D'), ALL_CATEGORIES, false); // blank allowed -> auto "Uncategorised"
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

  // Accountability rules (visual):
  //  - Category blank/"Uncategorised" on an Expense row -> flagged red.
  //    (onEdit in Rules.gs also writes the literal "Uncategorised" text;
  //    this formatting stays in place regardless of how a row was added.)
  //  - Work-scope row with no Receipt Link -> flagged red (ATO substantiation).
  var rules = [
    buildFlagRule(sheet.getRange('D2:D'), '=AND($C2="Expense",OR($D2="",$D2="Uncategorised"))'),
    buildFlagRule(sheet.getRange('G2:G'), '=AND($E2="Work",$G2="")')
  ];
  sheet.setConditionalFormatRules(rules);

  sheet.setColumnWidths(1, 10, 130);
}

function setupBudgetSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.BUDGET);
  setHeaderRow(sheet, ['Month', 'Category', 'Budgeted Amount']);

  sheet.getRange('A2:A').setNumberFormat(MONTH_FORMAT);
  sheet.getRange('C2:C').setNumberFormat(CURRENCY_FORMAT);
  applyListValidation(sheet.getRange('B2:B'), BUDGETABLE_CATEGORIES, true);

  sheet.setColumnWidths(1, 3, 160);
}

function setupSavingsGoalsSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.SAVINGS_GOALS);
  setHeaderRow(sheet, ['Goal Name', 'Target', 'Deadline', 'Current']);

  sheet.getRange('B2:B').setNumberFormat(CURRENCY_FORMAT);
  sheet.getRange('C2:C').setNumberFormat(DATE_FORMAT);
  sheet.getRange('D2:D').setNumberFormat(CURRENCY_FORMAT);

  // Current: sum of every Savings Contributions row for this goal. Same
  // SUMIF-with-array-criteria-under-ARRAYFORMULA pattern used elsewhere
  // (Transactions FY/Month, Reconciliation Tracked Net) so it covers every
  // goal row automatically.
  sheet.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",' +
    'SUMIF(\'Savings Contributions\'!$B$2:$B,A2:A,\'Savings Contributions\'!$C$2:$C)))'
  );
  protectFormulaColumn(sheet, 4, 1, 'Auto-summed: Current');

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

  // Tracked Net: Income minus Expense for the row's Month, computed from
  // Transactions. Wrapping SUMIFS in ARRAYFORMULA with a range (A2:A) as
  // one of the criteria evaluates it element-by-element, one result per row.
  sheet.getRange('C2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",' +
    'SUMIFS(Transactions!$B$2:$B,Transactions!$C$2:$C,"Income",Transactions!$I$2:$I,DATE(YEAR(A2:A),MONTH(A2:A),1))' +
    '-SUMIFS(Transactions!$B$2:$B,Transactions!$C$2:$C,"Expense",Transactions!$I$2:$I,DATE(YEAR(A2:A),MONTH(A2:A),1))))'
  );
  // Difference: blank until you've entered Actual Bank Net for that month,
  // so an un-reconciled month never shows as a false-positive flag.
  sheet.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(OR(A2:A="",B2:B=""),"",B2:B-C2:C))'
  );

  protectFormulaColumn(sheet, 3, 1, 'Auto-computed: Tracked Net');
  protectFormulaColumn(sheet, 4, 1, 'Auto-computed: Difference');

  var rules = [
    buildFlagRule(sheet.getRange('D2:D'), '=AND($D2<>"",ROUND($D2,2)<>0)')
  ];
  sheet.setConditionalFormatRules(rules);

  sheet.setColumnWidths(1, 4, 160);
}

function setupDashboardSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.DASHBOARD);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  sheet.getRange('A1').setValue('Budget Tracker Dashboard')
    .setFontWeight('bold').setFontSize(16);
  sheet.getRange('A2').setValue('Savings progress and receipt upload arrive in later build layers.')
    .setFontStyle('italic').setFontColor('#666666');

  // --- Month selector -----------------------------------------------
  sheet.getRange('A4').setValue('Select month (pick any date within it):').setFontWeight('bold');
  sheet.getRange('B4').setValue(new Date()).setNumberFormat(MONTH_FORMAT);
  applyDateValidation(sheet.getRange('B4'));
  sheet.getRange('A5').setValue('Financial year:').setFontWeight('bold');
  sheet.getRange('B5').setFormula(
    '="FY"&IF(MONTH(B4)>=7,YEAR(B4),YEAR(B4)-1)&"-"&RIGHT(IF(MONTH(B4)>=7,YEAR(B4)+1,YEAR(B4)),2)'
  );

  // Hidden-ish helper cell: Month normalised to the 1st, used by every
  // lookup formula below. Kept off in column N so the main view stays tidy.
  sheet.getRange('N1').setValue('Internal helper - do not edit').setFontColor('#999999').setFontSize(9);
  sheet.getRange('N4').setFormula('=DATE(YEAR(B4),MONTH(B4),1)').setNumberFormat(MONTH_FORMAT);
  setNamedRange(ss, 'DashMonth', sheet.getRange('N4'));

  // --- This month ------------------------------------------------------
  sheet.getRange('A7').setValue('THIS MONTH').setFontWeight('bold');
  sheet.getRange('A8').setValue('Income');
  sheet.getRange('B8').setFormula('=SUMIFS(Transactions!$B$2:$B,Transactions!$C$2:$C,"Income",Transactions!$I$2:$I,DashMonth)');
  sheet.getRange('A9').setValue('Expenses');
  sheet.getRange('B9').setFormula('=SUMIFS(Transactions!$B$2:$B,Transactions!$C$2:$C,"Expense",Transactions!$I$2:$I,DashMonth)');
  sheet.getRange('A10').setValue('Net').setFontWeight('bold');
  sheet.getRange('B10').setFormula('=B8-B9').setFontWeight('bold');
  sheet.getRange('B8:B10').setNumberFormat(CURRENCY_FORMAT);

  // --- Data quality & accountability -----------------------------------
  sheet.getRange('A12').setValue('DATA QUALITY & ACCOUNTABILITY').setFontWeight('bold');
  sheet.getRange('A13').setValue('Unassigned (Income − Budgeted), this month');
  sheet.getRange('B13').setFormula('=B8-B37');
  sheet.getRange('A14').setValue('Uncategorised expenses (all time)');
  sheet.getRange('B14').setFormula(
    '=COUNTIFS(Transactions!$C$2:$C,"Expense",Transactions!$D$2:$D,"")' +
    '+COUNTIFS(Transactions!$C$2:$C,"Expense",Transactions!$D$2:$D,"Uncategorised")'
  );
  sheet.getRange('A15').setValue('Work expenses missing a receipt (all time)');
  sheet.getRange('B15').setFormula('=COUNTIFS(Transactions!$E$2:$E,"Work",Transactions!$G$2:$G,"")');
  sheet.getRange('A16').setValue('Reconciliation months with a flagged difference');
  sheet.getRange('B16').setFormula('=COUNTIFS(Reconciliation!$D$2:$D,"<>0",Reconciliation!$D$2:$D,"<>")');
  sheet.getRange('B13').setNumberFormat(CURRENCY_FORMAT);

  var flagRules = [
    buildFlagRule(sheet.getRange('B13'), '=ROUND($B13,2)<>0'),
    buildFlagRule(sheet.getRange('B14'), '=$B14>0'),
    buildFlagRule(sheet.getRange('B15'), '=$B15>0'),
    buildFlagRule(sheet.getRange('B16'), '=$B16>0')
  ];

  // --- Spent vs budgeted by category, selected month --------------------
  sheet.getRange('A18').setValue('SPENT VS BUDGETED — SELECTED MONTH').setFontWeight('bold');
  var tableHeaderRow = 19;
  sheet.getRange(tableHeaderRow, 1, 1, 4).setValues([['Category', 'Budgeted', 'Spent', 'Remaining']])
    .setFontWeight('bold');

  var firstCatRow = tableHeaderRow + 1; // 20
  var lastCatRow = firstCatRow + EXPENSE_CATEGORIES.length - 1; // 36

  sheet.getRange(firstCatRow, 1, EXPENSE_CATEGORIES.length, 1)
    .setValues(EXPENSE_CATEGORIES.map(function (cat) { return [cat]; }));

  for (var i = 0; i < EXPENSE_CATEGORIES.length; i++) {
    var r = firstCatRow + i;
    sheet.getRange(r, 2).setFormula(
      '=SUMPRODUCT((Budget!$B$2:$B$1000=$A' + r + ')*(DATE(YEAR(Budget!$A$2:$A$1000),MONTH(Budget!$A$2:$A$1000),1)=DashMonth)*Budget!$C$2:$C$1000)'
    );
    sheet.getRange(r, 3).setFormula(
      '=SUMIFS(Transactions!$B$2:$B,Transactions!$C$2:$C,"Expense",Transactions!$D$2:$D,$A' + r + ',Transactions!$I$2:$I,DashMonth)'
    );
    sheet.getRange(r, 4).setFormula('=B' + r + '-C' + r);
  }
  sheet.getRange(firstCatRow, 2, EXPENSE_CATEGORIES.length, 3).setNumberFormat(CURRENCY_FORMAT);

  var totalsRow = lastCatRow + 1; // 37
  sheet.getRange(totalsRow, 1).setValue('Total').setFontWeight('bold');
  sheet.getRange(totalsRow, 2).setFormula('=SUM(B' + firstCatRow + ':B' + lastCatRow + ')').setFontWeight('bold');
  sheet.getRange(totalsRow, 3).setFormula('=SUM(C' + firstCatRow + ':C' + lastCatRow + ')').setFontWeight('bold');
  sheet.getRange(totalsRow, 4).setFormula('=B' + totalsRow + '-C' + totalsRow).setFontWeight('bold');
  sheet.getRange(totalsRow, 2, 1, 3).setNumberFormat(CURRENCY_FORMAT);

  flagRules.push(buildFlagRule(
    sheet.getRange(firstCatRow, 4, EXPENSE_CATEGORIES.length, 1),
    '=$D' + firstCatRow + '<0'
  ));

  // --- Top 5 categories by spend -----------------------------------------
  var topHeaderRow = totalsRow + 2; // 39
  sheet.getRange(topHeaderRow, 1).setValue('TOP 5 CATEGORIES BY SPEND — SELECTED MONTH').setFontWeight('bold');
  sheet.getRange(topHeaderRow + 1, 1, 1, 2).setValues([['Category', 'Spent']]).setFontWeight('bold');
  sheet.getRange(topHeaderRow + 2, 1).setFormula(
    '=IFERROR(ARRAY_CONSTRAIN(SORT({A' + firstCatRow + ':A' + lastCatRow + ',C' + firstCatRow + ':C' + lastCatRow + '},2,FALSE),5,2),"")'
  );
  sheet.getRange(topHeaderRow + 2, 2, 5, 1).setNumberFormat(CURRENCY_FORMAT);

  // --- Savings goals progress --------------------------------------------
  // Dynamic: driven by ARRAYFORMULA off a bounded 'Savings Goals' range, so
  // it spills to however many goals exist (up to GOAL_CAP) without a fixed
  // row count, but without ballooning the Dashboard sheet the way an
  // open-ended A2:A reference would (that spills ~1000 blank rows).
  var GOAL_CAP = 300;
  var savingsTitleRow = topHeaderRow + 2 + 5 + 1; // 47
  var savingsHeaderRow = savingsTitleRow + 1; // 48
  var firstGoalRow = savingsHeaderRow + 1; // 49
  var lastGoalRow = firstGoalRow + GOAL_CAP - 1;

  sheet.getRange(savingsTitleRow, 1).setValue('SAVINGS GOALS PROGRESS').setFontWeight('bold');
  sheet.getRange(savingsHeaderRow, 1, 1, 8).setValues([[
    'Goal Name', 'Target', 'Current', '% Complete', 'Deadline',
    'Avg Monthly Contribution', 'Projected Completion', 'On Track?'
  ]]).setFontWeight('bold');

  var g = firstGoalRow;
  var goalsName = '\'Savings Goals\'!$A$2:$A$' + (GOAL_CAP + 1);
  var goalsTarget = '\'Savings Goals\'!$B$2:$B$' + (GOAL_CAP + 1);
  var goalsDeadline = '\'Savings Goals\'!$C$2:$C$' + (GOAL_CAP + 1);
  var goalsCurrent = '\'Savings Goals\'!$D$2:$D$' + (GOAL_CAP + 1);
  var oRange = 'O' + g + ':O' + lastGoalRow;
  var pRange = 'P' + g + ':P' + lastGoalRow;
  var qRange = 'Q' + g + ':Q' + lastGoalRow;
  var fRange = 'F' + g + ':F' + lastGoalRow;

  // Hidden helpers, off in columns O/P/Q, one row per goal:
  //   O = date of that goal's first contribution (0 if none yet)
  //   P = whole months elapsed since then (0 if no contributions)
  //   Q = projected completion date (blank if reached / no contributions)
  sheet.getRange('O' + g).setFormula(
    '=ARRAYFORMULA(IF(' + goalsName + '="","",' +
    'MINIFS(\'Savings Contributions\'!$A$2:$A,\'Savings Contributions\'!$B$2:$B,' + goalsName + ')))'
  );
  sheet.getRange('P' + g).setFormula(
    '=ARRAYFORMULA(IF(' + goalsName + '="","",' +
    'IF(' + oRange + '=0,0,' +
    'IF((YEAR(TODAY())-YEAR(' + oRange + '))*12+(MONTH(TODAY())-MONTH(' + oRange + '))+1<1,1,' +
    '(YEAR(TODAY())-YEAR(' + oRange + '))*12+(MONTH(TODAY())-MONTH(' + oRange + '))+1))))'
  );
  sheet.getRange('A' + g).setFormula('=ARRAYFORMULA(IF(' + goalsName + '="","",' + goalsName + '))');
  sheet.getRange('B' + g).setFormula('=ARRAYFORMULA(IF(' + goalsName + '="","",' + goalsTarget + '))');
  sheet.getRange('C' + g).setFormula('=ARRAYFORMULA(IF(' + goalsName + '="","",' + goalsCurrent + '))');
  sheet.getRange('D' + g).setFormula(
    '=ARRAYFORMULA(IF(' + goalsName + '="","",IFERROR(' + goalsCurrent + '/' + goalsTarget + ',0)))'
  );
  sheet.getRange('E' + g).setFormula('=ARRAYFORMULA(IF(' + goalsName + '="","",' + goalsDeadline + '))');
  sheet.getRange('F' + g).setFormula(
    '=ARRAYFORMULA(IF(' + goalsName + '="","",IF(' + pRange + '=0,0,' + goalsCurrent + '/' + pRange + ')))'
  );
  sheet.getRange('Q' + g).setFormula(
    '=ARRAYFORMULA(IF(' + goalsName + '="","",' +
    'IF(' + goalsTarget + '-' + goalsCurrent + '<=0,"",' +
    'IF(' + fRange + '<=0,"",' +
    'EDATE(TODAY(),CEILING((' + goalsTarget + '-' + goalsCurrent + ')/' + fRange + ',1))))))'
  );
  sheet.getRange('G' + g).setFormula(
    '=ARRAYFORMULA(IF(' + goalsName + '="","",' +
    'IF(' + goalsTarget + '-' + goalsCurrent + '<=0,"Goal reached!",' +
    'IF(' + fRange + '<=0,"No contributions yet",TEXT(' + qRange + ',"mmm yyyy")))))'
  );
  sheet.getRange('H' + g).setFormula(
    '=ARRAYFORMULA(IF(' + goalsName + '="","",' +
    'IF(' + goalsDeadline + '="","",' +
    'IF(' + goalsTarget + '-' + goalsCurrent + '<=0,"On track",' +
    'IF(' + fRange + '<=0,"Behind",' +
    'IF(' + qRange + '<=' + goalsDeadline + ',"On track","Behind"))))))'
  );

  sheet.getRange('B' + g + ':C' + lastGoalRow).setNumberFormat(CURRENCY_FORMAT);
  sheet.getRange('D' + g + ':D' + lastGoalRow).setNumberFormat('0%');
  sheet.getRange('E' + g + ':E' + lastGoalRow).setNumberFormat(DATE_FORMAT);
  sheet.getRange('F' + g + ':F' + lastGoalRow).setNumberFormat(CURRENCY_FORMAT);

  flagRules.push(buildFlagRule(sheet.getRange('H' + g + ':H' + lastGoalRow), '=$H' + g + '="Behind"'));

  protectFormulaColumn(sheet, 15, 3, 'Internal helper: savings projection'); // O, P, Q

  sheet.setConditionalFormatRules(flagRules);

  sheet.setColumnWidths(1, 1, 320);
  sheet.setColumnWidths(2, 3, 130);
  sheet.setColumnWidths(5, 4, 150);
  sheet.hideColumns(14, 4); // N (DashMonth) .. Q (savings helpers)
}

function buildFlagRule(range, formula) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(FLAG_BG)
    .setRanges([range])
    .build();
}

function applyDateValidation(range) {
  var rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  range.setDataValidation(rule);
}

function setNamedRange(ss, name, range) {
  var existing = ss.getNamedRanges();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getName() === name) {
      existing[i].remove();
    }
  }
  ss.setNamedRange(name, range);
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
