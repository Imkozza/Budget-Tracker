/**
 * Menu entry point. Runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Budget Tracker')
    .addItem('Add Transaction...', 'showAddTransactionSidebar')
    .addItem('Seed Budget Month...', 'promptSeedBudgetMonth')
    .addSeparator()
    .addItem('Open Receipts Folder', 'showReceiptsFolderLink')
    .addItem('Export Tax Year...', 'promptTaxExport')
    .addSeparator()
    .addItem('Run Setup (first time / repair tabs)', 'setupWorkbook')
    .addToUi();
}

function showAddTransactionSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Add Transaction');
  SpreadsheetApp.getUi().showSidebar(html);
}
