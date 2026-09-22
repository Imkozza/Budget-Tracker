/**
 * Tax Export: every Work-scope transaction for a chosen financial year,
 * with receipt links and a total - written to the "Tax Export" tab and
 * saved as a PDF into the receipts Drive folder.
 */

function promptTaxExport() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Export Tax Year',
    'Enter the financial year\'s starting calendar year (e.g. 2025 for FY2025-26):',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var input = response.getResponseText().trim();
  if (!/^\d{4}$/.test(input)) {
    ui.alert('Could not parse "' + input + '". Enter a 4-digit year, e.g. 2025.');
    return;
  }

  var result = exportTaxYear(Number(input));

  var message = result.count === 0
    ? 'No Work-scope transactions found for ' + result.fyLabel + '.'
    : result.count + ' Work-scope transaction(s) for ' + result.fyLabel +
      ', totalling ' + formatAud(result.total) + '.' +
      (result.missingReceipts > 0
        ? '\n\n⚠ ' + result.missingReceipts + ' of these are missing a Receipt Link (highlighted in the Tax Export tab).'
        : '') +
      (result.pdfUrl ? '\n\nPDF saved: ' + result.pdfUrl : '\n\n(PDF export failed - see the Tax Export tab.)');

  ui.alert('Tax Export: ' + result.fyLabel, message, ui.ButtonSet.OK);
}

/**
 * @param {number} startYear Calendar year the financial year starts in (e.g. 2025 -> FY2025-26).
 * @return {Object} {fyLabel, count, total, missingReceipts, pdfUrl}
 */
function exportTaxYear(startYear) {
  var fyLabel = 'FY' + startYear + '-' + String(startYear + 1).slice(-2);

  var txnSheet = getSheet(SHEET_NAMES.TRANSACTIONS);
  var lastDataRow = getNextDataRow(txnSheet, 1) - 1;
  var rows = [];
  if (lastDataRow >= 2) {
    var data = txnSheet.getRange(2, 1, lastDataRow - 1, 10).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var scope = row[4];
      var fy = row[7];
      if (scope === 'Work' && fy === fyLabel) {
        rows.push({ date: row[0], amount: row[1], category: row[3], note: row[5], receiptLink: row[6] });
      }
    }
  }
  rows.sort(function (a, b) { return a.date - b.date; });

  var total = rows.reduce(function (sum, r) { return sum + Number(r.amount || 0); }, 0);
  var missingReceipts = rows.filter(function (r) { return !r.receiptLink; }).length;

  writeTaxExportSheet(fyLabel, rows, total);

  var pdfUrl = null;
  try {
    pdfUrl = saveTaxExportPdf(fyLabel);
  } catch (e) {
    // PDF export is a bonus on top of the tab; don't fail the whole export over it.
  }

  return { fyLabel: fyLabel, count: rows.length, total: total, missingReceipts: missingReceipts, pdfUrl: pdfUrl };
}

function writeTaxExportSheet(fyLabel, rows, total) {
  var ss = getSs();
  var sheet = getOrCreateSheet(ss, SHEET_NAMES.TAX_EXPORT);
  sheet.clear();
  sheet.clearConditionalFormatRules();

  sheet.getRange('A1').setValue('Work Expenses — ' + fyLabel).setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Generated ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'))
    .setFontStyle('italic').setFontColor('#666666');

  var headerRow = 4;
  sheet.getRange(headerRow, 1, 1, 5)
    .setValues([['Date', 'Amount', 'Category', 'Note', 'Receipt Link']])
    .setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
  sheet.setFrozenRows(headerRow);

  var firstDataRow = headerRow + 1;
  if (rows.length > 0) {
    var values = rows.map(function (r) {
      return [r.date, r.amount, r.category, r.note, r.receiptLink];
    });
    sheet.getRange(firstDataRow, 1, values.length, 5).setValues(values);
    sheet.getRange(firstDataRow, 1, values.length, 1).setNumberFormat(DATE_FORMAT);
    sheet.getRange(firstDataRow, 2, values.length, 1).setNumberFormat(CURRENCY_FORMAT);

    sheet.setConditionalFormatRules([
      buildFlagRule(sheet.getRange(firstDataRow, 5, values.length, 1), '=$E' + firstDataRow + '=""')
    ]);
  }

  var totalRow = firstDataRow + rows.length;
  sheet.getRange(totalRow, 1).setValue('Total').setFontWeight('bold');
  sheet.getRange(totalRow, 2).setValue(total).setFontWeight('bold').setNumberFormat(CURRENCY_FORMAT);

  sheet.setColumnWidths(1, 1, 100);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 260);

  moveSheetToEnd(ss, sheet);
}

function moveSheetToEnd(ss, sheet) {
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(ss.getSheets().length);
}

/**
 * Exports the Tax Export tab as a PDF and saves it into the receipts
 * Drive folder. Uses the spreadsheet's own export endpoint, authenticated
 * with the running script's OAuth token - a standard Apps Script pattern
 * (no external service involved).
 */
function saveTaxExportPdf(fyLabel) {
  var ss = getSs();
  var sheet = getSheet(SHEET_NAMES.TAX_EXPORT);
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf&gid=' + sheet.getSheetId() +
    '&size=A4&portrait=true&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=true&fzr=true';

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });

  var fileName = 'TaxExport_' + fyLabel + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.pdf';
  var blob = response.getBlob().setName(fileName);
  var folder = getOrCreateReceiptsFolder();
  var file = folder.createFile(blob);
  return file.getUrl();
}

function formatAud(amount) {
  return '$' + Number(amount).toFixed(2);
}
