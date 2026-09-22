/**
 * Budget tab helpers. The data model is "one row per category per month",
 * entered manually — this just removes the tedium of typing all category
 * rows for a new month by hand.
 */

/**
 * Prompts for a month (e.g. "2025-07") and inserts one Budget row per
 * budgetable category for that month, skipping categories that already
 * have a row for it. Budgeted Amount is left at 0 for you to fill in.
 */
function promptSeedBudgetMonth() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Seed Budget Month',
    'Enter the month to seed, as YYYY-MM (e.g. 2025-07):',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  var match = /^(\d{4})-(\d{2})$/.exec(input);
  if (!match) {
    ui.alert('Could not parse "' + input + '". Use the format YYYY-MM, e.g. 2025-07.');
    return;
  }
  var monthDate = new Date(Number(match[1]), Number(match[2]) - 1, 1);
  var added = seedBudgetMonth(monthDate);
  ui.alert(added === 0
    ? 'All budgetable categories already have a row for that month.'
    : 'Added ' + added + ' category row(s) for ' +
      Utilities.formatDate(monthDate, Session.getScriptTimeZone(), 'MMM yyyy') + '.');
}

/**
 * @param {Date} monthDate Any date; only year/month are used (normalised to the 1st).
 * @return {number} Number of rows added.
 */
function seedBudgetMonth(monthDate) {
  var sheet = getSheet(SHEET_NAMES.BUDGET);
  var target = monthStart(monthDate);

  var lastRow = getNextDataRow(sheet, 1) - 1; // last existing data row
  var existing = {};
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowMonth = data[i][0];
      if (rowMonth instanceof Date && monthStart(rowMonth).getTime() === target.getTime()) {
        existing[data[i][1]] = true;
      }
    }
  }

  var toAdd = BUDGETABLE_CATEGORIES.filter(function (cat) { return !existing[cat]; });
  if (toAdd.length === 0) return 0;

  var startRow = getNextDataRow(sheet, 1);
  var rows = toAdd.map(function (cat) { return [target, cat, 0]; });
  sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
  sheet.getRange(startRow, 1, rows.length, 1).setNumberFormat(MONTH_FORMAT);
  sheet.getRange(startRow, 3, rows.length, 1).setNumberFormat(CURRENCY_FORMAT);

  return rows.length;
}
