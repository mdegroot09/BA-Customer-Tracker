function RowInsert() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('4:4').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
};