/***********************************************************************************************************************
LIBRARY - Google Scripts utility functions and constants
***********************************************************************************************************************/


/*********************************************   CACHE    ******************************************/

// Workaround: App script doesnt work well with intializing global variables. 
// So declared globally, but initialized once in function
// getCache() to access cache object
var _cache;
function getCache(){           
  if ( ! _cache ) {
    _cache = new Cache();
    Logger.log("Initialized cache");
  }
  return _cache;
}
/***************************************************************************************************/



/*********************************************   SHEET    ******************************************/
/*
  Returns worksheet provided spreadsheet id, and worksheet's name
*/
function getSheet(spreadsheetId, worksheetName) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var worksheet = spreadsheet.getSheetByName(worksheetName);
  return worksheet;
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


/*
  Updates sheet based on values in a single-dimension array. 
  2-dimensional values not supported
 */
function updateSheet(sheet, startCol, startRow, values) {
  if ( values && values.length>0 ) {
    var rangeId = startCol + startRow;
    if (values.length > 1) rangeId = rangeId + ":" + startCol + (Number(startRow) + values.length - 1);
    var range = sheet.getRange(rangeId);
    var values2d = [];
    for (var i in values) {             // 2-d array mandated by Spreadsheet library
      values2d[i] = [values[i]];
    }
    range.setValues(values2d);
  }
}


/*
  Returns apt backup sheetname
*/
function getBackupSheetName(name){
  return name + "." + (new Date()).toString();
}


/*
  Makes a backup of provided sheet. Backup in same spreadsheet
*/
function backupSheet(sheet) {
  if (_MAX_SHEET_BACKUPS>0) {
    var name = sheet.getName();
    var spreadsheet = sheet.getParent();
    spreadsheet.insertSheet(getBackupSheetName(name), 1, {"template" : sheet});
    Logger.log("Created backup: " + getBackupSheetName(name));
  }
  removeOldBackups(sheet, _MAX_SHEET_BACKUPS);
}


var _MAX_SHEET_BACKUPS = 10;
/*
  Removes old backups of sheet reducing to number of max backups
*/
function removeOldBackups(sheet, maxBackups) {
  var name = sheet.getName();
  var spreadsheet = sheet.getParent();
  var allSheets = spreadsheet.getSheets();
  if ( allSheets.length > maxBackups + 1 ) {
    var countBackups = 0;
    var oldestBackup = undefined;
    for (var i in allSheets) {
      if ( allSheets[i].getName() == name ) continue;
      if (   allSheets[i].getName().indexOf( name ) == 0   ) {
        countBackups++; 
        if ( ! oldestBackup ) oldestBackup = allSheets[i];
        else {
          var oldDate = new Date( oldestBackup.getName().replace(name+".","") );
          var newDate = new Date( allSheets[i].getName().replace(name+".","") );
          if ( newDate < oldDate ) oldestBackup = allSheets[i];
        }
      }
    }
    if ( countBackups > _MAX_SHEET_BACKUPS ) {
      // Delete oldest backup
      var oldestBackupName = oldestBackup.getName();
      spreadsheet.setActiveSheet(oldestBackup);
      spreadsheet.deleteActiveSheet();
      Logger.log("Deleted  backup: " + oldestBackupName);
      //  Recurse till backups reduced
      if ( countBackups -1 > _MAX_SHEET_BACKUPS ) removeOldBackups(sheet, _MAX_SHEET_BACKUPS);
    }
  }
}
/***************************************************************************************************/



/*********************************************    URL     ******************************************/
/*
  Returns string content of provided url
*/
function fetchURL(url) {
  try {
    var xml = undefined;
    Logger.log("Attempting fetch " + url);
    var response = UrlFetchApp.fetch(url);
    Logger.log("response = " + response);
    //xml = response.getContentText();
    var blob = response.getBlob();
    Logger.log("blob = " + blob);
    var blobCopy = blob.copyBlob();
    Logger.log("blobCopy = " + blobCopy);
    xml = blob.getDataAsString()
    Logger.log("xml = " + xml);
  } catch(err) {
    error("Failed to fetch url: " + url);
    error(err.message);
  }
  Logger.log((xml ? "Fetched xml" : "Failed to fetch xml" ) + " : " + url);
  return xml;
}
/***************************************************************************************************/



/*********************************************   STRING   ******************************************/
function escapeSingleQuote(s){
  return s.replace(/'/g,"\\'");
}
/***************************************************************************************************/



/*********************************************   REGEX    ******************************************/
/*
  Updates all values in an array
  regex - valid regular expression
  replace - valid replacement expression, or plain string. Explained at https://developer.mozilla.org/en-US/docs/JavaScript/Reference/Global_Objects/String/replace
  Returns updated array on success, or undefined on error
*/
function updateValues(arr, regex, replace) {
  if ( ! regex || regex.length < 1 || ! isRegexValid(regex) ) {
    error("Invalid regex: " + regex);
    return undefined;
  } else if (! replace ) {
    error("Undefined replace string");
    return undefined;
  }
  else {
    var re = new RegExp(regex);
    var newArr = [];
    for (var i in arr) {
      var matches = arr[i].match(regex);
      if ( matches ) {
        newArr[i] = replace;              // Iterate, replacing $1, $2, placeholders
        for (var j=1; j<matches.length; j++) {
          newArr[i] = newArr[i].replace("$"+j, matches[j], "gi");
        }
//        Logger.log(newArr[i]);
      } else newArr[i] = undefined;
    }
    return newArr;
  }
}


/*
  true if valid, false otherwise
*/
function isRegexValid(regex){
  try {
    new RegExp(regex);
    return true;
  } catch (err) {
    error("Invalid regex '" + regex + "' " + err.message);
    return false;
  }
}


/***************************************************************************************************/



/*********************************************   ERROR    ******************************************/
/*
  Simply prints error to console
*/
function error(message) {
  Logger.log("ERROR: " + message);
}
/***************************************************************************************************/



/*********************************************    XML     ******************************************/
/* 
  Returns array (possibily empty) of all child elements matching xpath. 
  Very basic xpath only - no attributes or //.
*/
function getChildElements(element, xpath) {
  var children = [];
  if (!xpath || xpath.length<1) {
    error("Failed to find empty xpath '" + xpath + "'");       return children;
  } else {
    xpath = xpath.replace(/^\/+/,"").replace(/\/+$/,"");  // Remove starting and ending slashes
    var names = xpath.split("/");
    if ( names[0] != element.getName().getLocalName() ) {
      error("Failed to find XML element: " + names[0]);
    } else {
      // Logger.log("Found element: " + names[0]);
      if ( names.length == 1 ) {
        children.push(element);
      } else if (names.length > 1) {
        var childName = names[1];
        var childElements = element.getElements(childName);
        if ( !childElements || childElements.length < 1 ) {
          error("Failed to find element: " + childName);         return children;
        }
        var xpathChild = xpath.replace(names[0],"");
        for ( var i in childElements ) {
          var arr = getChildElements(childElements[i], xpathChild);
          children = children.concat(arr);
          // Dont cache. Recursion adds confusion
        }
      }
    }
  }
  return children;
}


/*
  Returns array of values for xpath from XMLElement. xpath can contain attributes. Returned array could be empty
*/
function getValues(element, xpath) {
  var values = [];
  var cacheKey = "getValues." + element.getName().getLocalName() + "." + xpath;
  if ( getCache().getItem(cacheKey) ) {
    values = getCache().getItem(cacheKey);
  } else {
    var elemPath = xpath.split("@")[0];
    var attrName = xpath.split("@")[1];
    var elements = getChildElements(element, elemPath);
    if ( elements.length == 0 ) {
      error("Failed to find xpath: " + xpath);
    } else {
      getCache().setItem(cacheKey, elements);
      for (var i in elements) {
        var value;
        if (attrName && attrName.length > 0) {
          var attr = elements[i].getAttribute(attrName);
          if (! attr) error("Failed to find attribute: " + attrName + ", in " + elements[i]);
          else  value = attr.getValue();
        } else {
          value = elements[i].getText();
        }
        if ( value ) values.push(value);
      }
    }
  }
  return values;
}
/***************************************************************************************************/



/*********************************************    TIME    ******************************************/
var START_TIME = (new Date()).getTime();
var MAX_RUNTIME_SECONDS = 28*60;  //28 mins


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