/**
 * Central configuration for the Budget Tracker workbook.
 * AU / Queensland: currency AUD, financial year 1 Jul - 30 Jun, no DST.
 */

var SHEET_NAMES = {
  TRANSACTIONS: 'Transactions',
  BUDGET: 'Budget',
  SAVINGS_GOALS: 'Savings Goals',
  SAVINGS_CONTRIBUTIONS: 'Savings Contributions',
  RECONCILIATION: 'Reconciliation',
  DASHBOARD: 'Dashboard',
  TAX_EXPORT: 'Tax Export'
};

// Sheet tab order, left to right.
var SHEET_ORDER = [
  SHEET_NAMES.DASHBOARD,
  SHEET_NAMES.TRANSACTIONS,
  SHEET_NAMES.BUDGET,
  SHEET_NAMES.SAVINGS_GOALS,
  SHEET_NAMES.SAVINGS_CONTRIBUTIONS,
  SHEET_NAMES.RECONCILIATION
];

var DIRECTIONS = ['Income', 'Expense'];
var SCOPES = ['Personal', 'Work'];
var RECONCILED_OPTIONS = ['Y', 'N'];

var EXPENSE_CATEGORIES = [
  'Housing',
  'Utilities',
  'Groceries',
  'Transport',
  'Insurance',
  'Health & Medical',
  'Personal Care',
  'Entertainment',
  'Dining Out',
  'Subscriptions',
  'Education',
  'Work Expenses',
  'Gifts & Donations',
  'Travel',
  'Debt Repayment',
  'Miscellaneous',
  'Uncategorised'
];

var INCOME_CATEGORIES = ['Salary', 'Bonus', 'Interest', 'Other Income'];

var ALL_CATEGORIES = EXPENSE_CATEGORIES.concat(INCOME_CATEGORIES);

var UNCATEGORISED_LABEL = 'Uncategorised';

// Categories you can actually set a Budgeted Amount against — excludes
// "Uncategorised" itself, since the goal is to have nothing land there.
var BUDGETABLE_CATEGORIES = EXPENSE_CATEGORIES.filter(function (c) {
  return c !== UNCATEGORISED_LABEL;
});

// Script Properties keys.
var PROPS_RECEIPTS_FOLDER_ID = 'RECEIPTS_FOLDER_ID';

// The one dedicated Drive folder the spec calls for: receipt images AND
// generated Tax Export PDFs both live here.
var RECEIPTS_FOLDER_NAME = 'Budget Tracker Receipts';

var MAX_RECEIPT_FILE_BYTES = 10 * 1024 * 1024; // 10 MB

var CURRENCY_FORMAT = '$#,##0.00';
var DATE_FORMAT = 'yyyy-mm-dd';
var MONTH_FORMAT = 'mmm yyyy';

var HEADER_BG = '#1c4587';
var HEADER_FG = '#ffffff';
var FLAG_BG = '#f4cccc';
