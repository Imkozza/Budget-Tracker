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

// Script Properties keys (used from Layer 4 onward).
var PROPS_RECEIPTS_FOLDER_ID = 'RECEIPTS_FOLDER_ID';

var CURRENCY_FORMAT = '$#,##0.00';
var DATE_FORMAT = 'yyyy-mm-dd';
var MONTH_FORMAT = 'mmm yyyy';

var HEADER_BG = '#1c4587';
var HEADER_FG = '#ffffff';
var FLAG_BG = '#f4cccc';
