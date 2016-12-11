function onEdit() {
  // Copied from: https://productforums.google.com/d/topic/docs/ehoCZjFPBao/discussion
  var sheetNameToWatch = "New";
  var columnNumberToWatch = 12; // column A = 1, B = 2, etc.
  var valueToWatch = "x";
  var targetSheetToMoveTheRowTo = "Upcoming";
  var rangeToSortOnSheetOnTargetSheet = "A2:T";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveCell();
  
  if (sheet.getName() == sheetNameToWatch && range.getColumn() == columnNumberToWatch && range.getValue() == valueToWatch) {
    var targetSheet = ss.getSheetByName(targetSheetToMoveTheRowTo);
    var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
    sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).moveTo(targetRange);
    sheet.deleteRow(range.getRow());
  }
  
  // Do sort and archive
  sortAndArchive();
}

function myFunction() {
  var s = SpreadsheetApp.getActiveSheet();
  var values = s.getDataRange().getValues();
  nextLine: for( var i = values.length-1; i >=0; i-- ) {
    for( var j = 0; j < values[i].length; j++ )
      if( values[i][j] != "" )
        continue nextLine;
    s.deleteRow(i+1);
  }
  //I iterate it backwards on purpose, so I do not have to calculate the indexes after a removal
}												

function sortAndArchive() {
  
  // Important parameters
  var dateColumn = 8;
  var dateIndex = dateColumn - 1;
  var olderInTop = true;
  
  //
  var today = new Date();
  today.setHours(0,0,0,0);
 
  // Sheets.
  var app = SpreadsheetApp.getActiveSpreadsheet();
  var upcomming = app.getSheetByName("Upcoming");
  var archive = app.getSheetByName("Archive");
 
  // Sort range A2:T by dateColumn.
  var allColumns = upcomming.getRange("A2:T");
  allColumns.sort({column: dateColumn, ascending: olderInTop});
  
  
  // Do color coding.
  var colorRange = upcomming.getRange("H1:H999");
  var cellsToColor = colorRange.getValues();
  
  var todayStart = new Date(today);
  todayStart.setHours(0, 0, 0, 0);
  var todayEnd = new Date(today);
  todayEnd.setHours(23, 59, 59, 999);

  var tomorrowStart = new Date(todayStart);
  tomorrowStart.setDate(todayStart.getDate() + 1);
  var tomorrowEnd = new Date(todayEnd);
  tomorrowEnd.setDate(todayEnd.getDate() + 1);
  
  for (var rowI=1; rowI<cellsToColor.length; ++rowI) {  
    var cellDate = new Date(cellsToColor[rowI][0]);
    
    if (cellDate != "" && cellDate >= todayStart && cellDate <= todayEnd) {
      setColor(rowI, dateIndex, upcomming, "#b7e1cd");
    } else if (cellDate != "" && cellDate >= tomorrowStart && cellDate <= tomorrowEnd) {
      setColor(rowI, dateIndex, upcomming, "#fff2cc");
    } else {
      setColor(rowI, dateIndex, upcomming, "#ffffff");
    }
    
  }
  
  // Run this foreveer and exit when script didn't archive anything.
  while(true) {
    
    // Assume that we didn't archive anything yet.
    var archivedSomething = false;
    
    var rows = allColumns.getValues();
    
    // Scan every row in range.
    for (var rowI = 0; rowI < rows.length; rowI++) {
      
      var interviewDate = rows[rowI][dateIndex];      
      
      // Check if interviewDate not empty AND interview date is in the past.
      if (interviewDate != "" && interviewDate < today) {
       
        // Move row rowI to archive.
        var sourceRange = upcomming.getRange(rowI + 2, 1, 1, upcomming.getLastColumn());
        var targetRange = archive.getRange(archive.getLastRow() + 1, 1);
        sourceRange.copyTo(targetRange);
        upcomming.deleteRow(rowI + 2);
        
        // Tell script that it archived something and stop scanning.
        archivedSomething = true;
        break;
      
      }
    }

    // If didn't archive anything, break.
    if (!archivedSomething) {
      break;
    }
  }

}

function setColor(rowI, colI, sheet, colorCode) {
  var range = sheet.getRange(rowI + 1, colI + 1);
  range.setBackgroundColor(colorCode);
}
