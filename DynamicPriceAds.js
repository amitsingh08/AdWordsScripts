/****************************************************************************************************
 AdWords Scripts, code + library, for AdWords-automation
 ****************************************************************************************************/

/****************************************************************************************************
 CONSTANTS
 ****************************************************************************************************/

var DYNAMIC_VALUE_SHEET_ID = "0AgtmhFLLc2KadG9VMHZIb2hIeVY5emlVRHI2NXMyTGc";

var sheetName = "DynamicValues";

var KEY_VALUE_SHEET_NAME = "key=value";

var MAX_RUNTIME_SECONDS = 28*60;  //28 mins

var START_TIME = (new Date()).getTime();


/****************************************************************************************************
 MAIN
 ****************************************************************************************************/

function main() {
  Logger.log("START");
  var values = getSheetValues(DYNAMIC_VALUE_SHEET_ID, sheetName);
  var lastRowProcessed = getValue("LastRowProcessed");
  var i = (lastRowProcessed && lastRowProcessed<values.length) ? lastRowProcessed : 1;
  Logger.log("Start from row " + (i+1));
  while (i<values.length) {
    var camp = escapeSingleQuote(values[i][1].trim());
    var adGroup = escapeSingleQuote(values[i][2].trim());
    var param1 = values[i][3];
    var param2 = values[i][4];
    if (checkValues(camp, adGroup, param1, param2)) {
      updateKeywordParams(camp,adGroup,param1,param2);
      Logger.log('Row processed: ' + (i+1) );
    } else {
      Logger.log('Failed processing row: ' + (i+1) );
    }
    if ( isTimeout() ) {
      lastRowProcessed = i+1;
      Logger.log("Stop premature at row " + lastRowProcessed);
      break;
    } else {
      i++;
    }
  }
  if ( i==values.length ) {
    Logger.log("All rows processed");
    lastRowProcessed = values.length;
  }
  getKV()["LastRowProcessed"] = lastRowProcessed;
  preExit();
  Logger.log("END");
}

/****************************************************************************************************
 FUNCTIONS
 ****************************************************************************************************/

function test() {
  var kv = getKV();
  kv["EndLastTime"] = new Date();
  saveKV();
}

/*
 Returns true only if values are valid
 */
function checkValues(camp, adGroup, param1, param2) {
  if ( camp.length < 1 || adGroup.length < 1 || (param1 && isNaN(param1)) || (param2 && isNaN(param2)) ) {
    Logger.log('Invalid values: ' + camp + ',' + adGroup + ',' + param1 + ',' + param2);
    return false;
  }
  else return true;
}

function updateKeywordParams(camp, adGroup, param1, param2) {
  var keywordIterator = getKeywordIterator(camp, adGroup);
  var count = 0;
  while (keywordIterator.hasNext()){
    var keyword = keywordIterator.next();
    if (param1) keyword.setAdParam(1, param1);
    if (param2) keyword.setAdParam(2, param2);
    count++;
  }
  if ( count > 0 ) Logger.log('Updated : "' + camp + '" > "' + adGroup + '" : ' + param1 + ', ' + param2 + ' : keywords: ' + count);
}

/*
 Called just-before exit
 */
function preExit(){
  getKV()["StartLastTime"] = new Date(START_TIME);
  getKV()["EndLastTime"] = new Date();
  saveKV();
}

/****************************************************************************************************
 LIBRARY
 ****************************************************************************************************/

/*
 id - spreadsheet id
 sheetName (optional) - name of sheet
 If specified sheetName doesnt exist, returns null. If sheetName not specified, active sheet returned
 */
function getSheet(id, sheetName){
  var spreadsheet = SpreadsheetApp.openById(id);
  Logger.log('Opened spreadsheet: ' + spreadsheet.getName());
  return sheetName ? spreadsheet.getSheetByName(sheetName) : spreadsheet.getActiveSheet();
}

/*
 Get 2 dimensional array of values based on sheet id passed. sheetName is optional
 */
function getSheetValues(id, sheetName) {
  var values = undefined;
  var sheet = getSheet(id, sheetName);
  if ( ! sheet ) {
    Logger.log("Sheet not found - id:" + id + " , sheetName:" + sheetName);
  } else {
    values = getValuesFromSheet(sheet);
  }
  return values;
}

/*
 Get's all values from passed sheet (worksheet, not spreadsheet which contains worksheets).
 Returns 2-dimensional array
 */
function getValuesFromSheet(sheet) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  Logger.log('Read sheet: ' + sheet.getName() + ' rows=' + values.length + ' cols=' + values[0].length);
  return values;
}

/*
 Overwrites values in a sheet, to those specified in 2 dimensional array.
 array2dimension must be a 2 dimensional array, with each row containing same number of entries
 */
function setSheetValues(sheet, array2dimension) {
  var rows = array2dimension.length;
  var cols = array2dimension[0].length;
  //Clear old values
  var range = sheet.getRange(1,1,rows,cols);
  range.setValues(array2dimension);
}

function getKeywordIterator(camp, adGroup) {
  var keywordIterator = AdWordsApp.keywords().
    withCondition("CampaignName = '" + camp + "'").
    withCondition("AdGroupName = '" + adGroup + "'").
    get();
  return keywordIterator;
}

/*
 Gets key-values from key=value sheet. Requires spreadheet and sheet read access
 */
function getKV() {
  if ( ! kv ) {
    var arr = getSheetValues(DYNAMIC_VALUE_SHEET_ID, KEY_VALUE_SHEET_NAME);
    if ( ! arr ) {
      Logger.log("Failed to read key/values from: " + DYNAMIC_VALUE_SHEET_ID + "." + KEY_VALUE_SHEET_NAME);
    }
    else {
      kv = new Object();
      for (var i=0; i<arr.length; i++) {
        var key = arr[i][0];
        if (key) {
          var value = arr[i][1];
          kv[key] = value;
        }
      }
      Logger.log("Initialized key-values from sheet: " + KEY_VALUE_SHEET_NAME);
    }
  }
  return kv;
}

var kv = undefined;

/*
 Saves key-values in key=value sheet
 */
function saveKV(){
  var array2dimension = [];
  var i=0;
  for ( var key in kv ) {
    var value = kv[key];
    array2dimension.push([key,value]);
    i++;
  }
  var sheet = getSheet(DYNAMIC_VALUE_SHEET_ID, KEY_VALUE_SHEET_NAME);
  if (sheet) {
    setSheetValues(sheet,array2dimension);
    Logger.log("Saved key-values");
  }
}


/*
  Returns value for key - or undefined, if key doesnt exist
*/
function getValue(key){
  var o = getKV();
  return o[key] ? o[key] : undefined;
}


function escapeSingleQuote(s){
  return s.replace(/'/g,"\\'");
}


/*
 Returns seconds left before timeout
 */
function timeRemaining() {
  return MAX_RUNTIME_SECONDS - timeRun();
}

// Seconds, the script has run
function timeRun() {
  return ((new Date()).getTime() - START_TIME)/1000;
}

//  Returns boolean true, only if script run-time is over. Flase otherwise
function isTimeout() {
  return timeRemaining() <= 0 ? true : false;
}