/**
 * Accountability rule: every Expense must have a Category. Blank is
 * treated as "Uncategorised" — this simple trigger writes that literal
 * text back into the cell whenever someone edits Transactions directly
 * (the sidebar path is handled server-side in Transactions.gs). The red
 * highlight itself comes from the conditional format rule set up in
 * Setup.gs, so it applies no matter how the row was created.
 */
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.TRANSACTIONS) return;

  var startRow = Math.max(range.getRow(), 2);
  var endRow = range.getLastRow();
  if (endRow < startRow) return;

  var numRows = endRow - startRow + 1;
  var directions = sheet.getRange(startRow, 3, numRows, 1).getValues();
  var categories = sheet.getRange(startRow, 4, numRows, 1).getValues();
  var changed = false;

  for (var i = 0; i < numRows; i++) {
    var direction = directions[i][0];
    var category = categories[i][0];
    if (direction === 'Expense' && (category === '' || category === null)) {
      categories[i][0] = UNCATEGORISED_LABEL;
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(startRow, 4, numRows, 1).setValues(categories);
  }
}
