/************************************************************************************
  Spreadsheet library providing opening, querying, updating, data
  Author: amitsingh@google.com

  Example Code Below:
  queryTest() - how to run a query against a spreadsheet
************************************************************************************/


/************************************************************************************
  MAIN
************************************************************************************/
// var url = 'https://docs.google.com/a/google.com/spreadsheet/ccc?key=0AguYBtcInoaldHZHUHVzbkpNcTNMZW80OHJyUTNiUXc#gid=0'; //Amit's sheet
var url = 'https://docs.google.com/a/google.com/spreadsheet/ccc?key=0AuxgPizDRGWKdHI5RDQ0dk82MWtQS3pwU1ZJN3ctY2c&usp=sharing#gid=8';
// var worksheetName = 'Account Stats';
var worksheetName = 'RmktListSize';
//var sql = "SELECT CID, UserListId, ListSize, Date, StatsDate where StatsDate <= date '2013-12-11' ORDER BY StatsDate";
var sql = "SELECT CID, UserListId, datevalid, sum(listsize) WHERE datevalid >= date '2013-08-22' AND DateValid <= date '2013-12-11' GROUP BY CID, UserListId, DateValid ORDER BY CID, UserListId, DateValid" //Yan's query
function queryTest() {
  var array2d = querySpreadsheet(url, worksheetName, sql);
  if ( array2d && array2d.length >=1 ) {
    for ( var i in array2d ) {
      Logger.log('CID=' + array2d[i][0] + ' | UserListId=' + array2d[i][1] + ' | StatsDate=' + array2d[i][2] + ' | ListSize=' + array2d[i][3]);
    }
  } else {
    Logger.log('No data from query. ' + url + ' | ' + worksheetName + ' | ' + sql);
  }
}

/************************************************************************************
  CONSTANTS
************************************************************************************/
var QUERY_RESULTS_SHEET = 'QueryResults';
var MAX_COLS_IN_SHEET = 256;

/************************************************************************************
  FUNCTIONS
************************************************************************************/

//   Returns true, or false
function valid(sheet){
  return sheet ?  true : false;
}

//   Returns spreadsheet from URL, or undefined on error
function getSpreadsheet(url) {
  var spreadsheet = undefined;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(url);
  } catch (ex) {
    error("Failed to get sheet: " + url + "\n" + ex);
  }
  return spreadsheet;
}

//   Returns 2darray after querying worksheet
//   sql parameter uses Spreadsheet QUERY() formula syntax
//   2darray empty when no result. Doesnt check if query valid.
function queryWorksheet(worksheet, sql) {
  var array2d = [0][0];
  Logger.log(worksheet.getName() + ' | ' + sql);
  
  //  Get Column names. Assume only 1 header row
  var headerRow = worksheet.getRange('1:1');
  var columnNames = headerRow.getValues()[0];
  if ( ! columnNames || columnNames.length < 1 ) {
    error('Failed getting column names in worksheet');
    return array2d;
  }
  
  //  'date' is reserved word, and can also be column-name. Eg "Date > date '2013-1-1'". 
  //   Protect 'date' keyword by renaming it to '[d a t e]'
  sql = sql.replace(/date( +['"])/gi,'[d a t e]$1');

  //  Update sql replacing column names with column index
  for (var i=0; i<columnNames.length; i++) {
    var colNotation = columnNotation(i+1);
    var regex = new RegExp('\\b' + columnNames[i] + '\\b', 'igm');
    sql = sql.replace(regex, colNotation);
  }

  //  Unprotect date keyword i.e. reverse '[d a t e]' to 'date'
  sql = sql.replace(/\[d a t e\]/g,'date');
  
  //  Execute sql
  var spreadsheet = worksheet.getParent();
  var resultsSheet = spreadsheet.getSheetByName(QUERY_RESULTS_SHEET);
  if ( ! resultsSheet ) resultsSheet = spreadsheet.insertSheet(QUERY_RESULTS_SHEET);
  resultsSheet.clear();
  var cell = resultsSheet.getRange('A1');
  var dataQueryRange = "'" + worksheet.getName() + "'!A:" + columnNotation(MAX_COLS_IN_SHEET);
  var formula = '=query(' + dataQueryRange + ',\"' + sql + '\", -1)';
  // Logger.log('Cell formula: ' + formula);
  cell.setFormula(formula);
  
  //  Fetch results
  array2d = resultsSheet.getDataRange().getValues();
  array2d.shift();  // Header row remove
  if ( array2d.length == 0 ) Logger.log('No results from query');
  else Logger.log('Query result: rows=' + array2d.length + ' cols=' + array2d[0].length);
  return array2d;
}


//   Returns 2darray after querying worksheet. Empty error can also be returned
function querySpreadsheet(spreadsheetURL, worksheetName, sql) {
  var array2d = [0][0];
  var spreadsheet = getSpreadsheet(url);
  if ( spreadsheet ) {
    var worksheet = spreadsheet.getSheetByName(worksheetName);
    if ( ! worksheet ) {
      error("Spreadsheet, failed getting worksheet: " + worksheetName); 
    } else {
      array2d = queryWorksheet(worksheet, sql);
    }
  }
  return array2d;
}


//  Returns column notation for passed number
//  1 return A, 26 returns Z, 27 returns AA, 53 returns BA, 702 return ZZ, 703 returns AAA
function columnNotation(number){
  var letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  var remainder = (number-1) % 26;
  var quotient = Math.floor( (number-1)/26 );
  // Recurse, when quotient present
  var columnName = quotient >= 1 ? columnNotation(quotient) + letters.charAt(remainder) : letters.charAt(remainder);
  return columnName;
}
                   

//   Handles error message. Returns nothing.
function error(log) {
  Logger.log("ERROR: " + log);
}
