function cleaner() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var tb1 = ss.insertSheet();
  tb1.setName("Tabellenblatt1")
  ss.moveActiveSheet(ss.getNumSheets());

  for (i = 0; i < sheets.length; i++) {
    ss.deleteSheet(sheets[i]);
  }
}
