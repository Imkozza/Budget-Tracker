/**
 * Server-side functions backing the "Add Transaction" sidebar.
 */

function getTransactionFormOptions() {
  return {
    directions: DIRECTIONS,
    categories: ALL_CATEGORIES,
    scopes: SCOPES,
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

/**
 * Adds one transaction row. Financial Year / Month are left blank because
 * they are auto-computed for the whole column by the ARRAYFORMULA set up
 * in Setup.gs. Reconciled defaults to "N".
 *
 * @param {Object} form {date, amount, direction, category, scope, note, receiptLink}
 * @return {Object} {row, message}
 */
function addTransaction(form) {
  var errors = validateTransactionForm(form);
  if (errors.length) {
    throw new Error(errors.join(' '));
  }

  var sheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var row = getNextDataRow(sheet, 1); // column A = Date is the reliable "key" column
  var dateValue = parseFormDate(form.date);
  var categoryInput = (form.category || '').trim();
  var wasBlank = !categoryInput && form.direction === 'Expense';
  // onEdit (Rules.gs) only fires for user-driven sheet edits, not writes
  // made from server code like this one, so apply the same "blank expense
  // -> Uncategorised" rule here.
  var category = wasBlank ? UNCATEGORISED_LABEL : categoryInput;

  sheet.getRange(row, 1, 1, 7).setValues([[
    dateValue,
    Number(form.amount),
    form.direction,
    category,
    form.scope,
    form.note || '',
    form.receiptLink || ''
  ]]);
  sheet.getRange(row, 10).setValue('N'); // Reconciled

  SpreadsheetApp.flush();

  var warning = wasBlank
    ? ' Category was left blank — set to "Uncategorised" and flagged.'
    : '';

  return {
    row: row,
    message: 'Transaction added on row ' + row + '.' + warning
  };
}

function validateTransactionForm(form) {
  var errors = [];
  if (!form || !form.date) {
    errors.push('Date is required.');
  }
  var amount = Number(form && form.amount);
  if (!form || form.amount === '' || isNaN(amount) || amount <= 0) {
    errors.push('Amount must be a positive number.');
  }
  if (!form || DIRECTIONS.indexOf(form.direction) === -1) {
    errors.push('Direction must be Income or Expense.');
  }
  if (!form || SCOPES.indexOf(form.scope) === -1) {
    errors.push('Scope must be Personal or Work.');
  }
  return errors;
}
