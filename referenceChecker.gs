//global variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var s0;
var s3 = ss.getSheetByName("Extensions")
var s4 = ss.getSheetByName("Rules");
var s5 = ss.getSheetByName("DataElements");
var s8 = ss.getSheetByName("DEandCV");

//get values from column A in DEandCV sheet
var s8colA = s8.getRange("A:A").getValues().flat();
var lastRow = s8colA.filter(String).length;

function handler() {
  setReferencesSheet();
  getNoteValues();
  getDataElementReferences();
  getRulesReferences();
  getExtensionsReferences();
  formatter();
}

//rename first sheet and add header row
function setReferencesSheet() {
  s0 = ss.getSheets()[0]
  s0.setName("References");
  s0.getRange("A1").setValue("Data Element Name");
  s0.getRange("B1").setValue("Data Element References");
  s0.getRange("C1").setValue("Rule References");
  s0.getRange("D1").setValue("Extension References");
  s0.getRange(1, 1, 1, s0.getLastColumn()).setFontWeight('bold');
}

//replace values in column H of data elements sheet with full code from notes
function getNoteValues() {
  var s5colH = s5.getRange("H:H");
  var notes = s5colH.getNotes();

  //override cell values in column H with note values
  for(var i = 0; i < notes.length; i++) {
    for (var j = 0; j < notes[0].length; j++) {
      // If note is not empty
      if(notes[i][j]) {
        var note = notes[i][j];
        var cell = s5colH.getCell(i+1,j+1);
        cell.setValue(note);
      }
    }
  }
}

//get all data element references
function getDataElementReferences() {

  //get values from columns A and K in DataElements sheet
  var s5colA = s5.getRange("A:A").getValues().flat();
  var s5colH = s5.getRange("H:H").getValues().flat();
  
  for (var i = 1; i < lastRow; i++) {
    var searchString = s8colA[i];
    var searchPatterns = [
      "%" + searchString + "%",
      "_satellite.getVar('" + searchString + "')",
      '_satellite.getVar("' + searchString + '")'
    ];

    //filter values in DataElements sheet by searchPatterns
    var filtered = s5colA.filter(function(value, index) {
      return searchPatterns.some(function(pattern) {
        return s5colH[index].includes(pattern);
      });
    });

    //join results with newline char, or return empty string
    var result = filtered.length > 0 ? filtered.join('\n') : "";

    //add DE names and results to AllReferences sheet
    s0.getRange(i + 1, 1).setValue(searchString);
    s0.getRange(i + 1, 2).setValue(result);
  }
}

//get all rules references
function getRulesReferences() {

  //get values from columns A and F in Rules sheet
  var s4colA = s4.getRange("A:A").getValues().flat();
  var s4colF = s4.getRange("F:F").getValues().flat();

  for (var i = 1; i < lastRow; i++) {
    var searchString = s8colA[i];
    if (!searchString) {
      continue;
    }

    //filter values in Rules sheet by searchString
    var filtered = s4colA.filter(function(value, index) {
      return s4colF[index].includes(searchString);
    });

    //join results with newline char, or return empty string
    var result = filtered.length > 0 ? filtered.join('\n') : "";

    //add results to AllReferences sheet
    s0.getRange(i + 1, 3).setValue(result);
  }
}

//get all extension references
function getExtensionsReferences() {

  //get values from columns A and F in Extensions sheet
  var s3colA = s3.getRange("A:A").getValues().flat();
  var s3colB = s3.getRange("B:B").getValues().flat();

  for (var i = 1; i < lastRow; i++) {
    var searchString = s8colA[i];
    if (!searchString) {
      continue;
    }

    //filter values in Rules sheet by searchString
    var filtered = s3colA.filter(function(value, index) {
      return s3colB[index].includes(searchString);
    });

    //join results with newline char, or return empty string
    var result = filtered.length > 0 ? filtered.join('\n') : "";

    //add results to AllReferences sheet
    s0.getRange(i + 1, 4).setValue(result);
  }
}

//format AllReferences sheet and higlight unreferenced data elements
function formatter() {
  var columns = [1, 2, 3, 4];
  var dataRange = s0.getDataRange()
  var values = dataRange.getValues();

  // Set vertical alignment to top and resize to fit text
  columns.forEach(function(column) {
    s0.getRange(1, column, s0.getMaxRows(), 1).setVerticalAlignment("top");
    s0.autoResizeColumn(column);
  });

  //freeze first row
  s0.setFrozenRows(1)

  //set background to yellow if columns B and C are empty
  for (var i = 1; i < values.length; i++) {
    var cellB = values[i][1];
    var cellC = values[i][2];
    var cellD = values[i][3];

    if (cellB === "" && cellC === "" && cellD ==="") {
      var range = s0.getRange(i + 1, 1, 1, s0.getLastColumn()); // get entire row
      range.setBackground('yellow');
    }
  }
}
