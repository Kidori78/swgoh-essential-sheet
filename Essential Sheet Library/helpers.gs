/*
Functions found here
 - deleteUnusedCells(sheet, addRow, addCol)
 - getLocalizationLang(language)
 - translate(message, targetLanguage)
 - sortSheets(targetSheet,columns)
 - deleteRows(targetSheet)
 - deleteEntries(targetSheet,players,allyCol,cleanup)
 - appendData(targetSheet, data)
 - getOmicronArea(mode)
 - getModSearchID(mod)
 - getColumnLetter(col)
*/


/**
 * Used to delete unused cells to save on cell count
 * @param {object} sheet - The sheet object to adjust
 * @param {integer} addRow - The number of unused rows to leave
 * @param {integer} addCol - The number of unused columns to leave
 */
function deleteUnusedCells(sheet,addCol = 0){
  SpreadsheetApp.flush();
  let maxCols = sheet.getMaxColumns();
  let curCols = sheet.getLastColumn() + addCol;
  if(maxCols !== curCols){
    sheet.deleteColumns((curCols + 1), (maxCols - curCols));
  }
  deleteRows(sheet)
}

/**************************************
 * Converts language into the corresponding iso language code
 * @param {string} language - The language name to convert to iso
 * @returns {string} The iso format for the language
 */
function getLocalizationLang(language){
  switch(language){
    case "中文(简体)":  //Chinese Simplified
      return "CHS_CN";
    case "中國傳統的": //Chinese Traditional
      return "CHT_CN";
    case "Français":  //French
      return "FRE_FR";
    case "Deutsch":   //German
      return "GER_DE";
    case "Indonesia": //Indonesian
      return "IND_ID";
    case "Italiano":  //Italian
      return "ITA_IT";
    case "日本語":    //Japanese
      return "JPN_JP";
    case "한국인": //Korean
      return "KOR_KR";
    case "Português (Brasil)":  //Potrugese (Brazil)
      return "POR_BR"
    case "Русский":   //Russian
      return "RUS_RU";
    case "Español":   //Spanish
      return "SPA_XM";
    case "แบบไทย": //Thai
      return "THA_TH";
    case "Türkçe":    //Turkish
      return "TUR_TR";
    default:          //English
      return "ENG_US";
  }
}

/**************************************
 * Translates the provided message into the provided language
 * @param {string} message - The message to convert
 * @param {string} language - The language to convert it into
 */
function translate(message, language){
  if(language.length !== 6 && language.indexOf("_") !== 3){
    language = getLocalizationLang(language);
  }
  switch(language){
    case "CHS_CN":
      return LanguageApp.translate(message,'en','zh');
    case "CHT_CN":
      return LanguageApp.translate(message,'en','zh-TW');
    case "FRE_FR":
      return LanguageApp.translate(message,'en','fr');
    case "GER_DE":
      return LanguageApp.translate(message,'en','de');
    case "IND_ID":
      return LanguageApp.translate(message,'en','id');
    case "ITA_IT":
      return LanguageApp.translate(message,'en','it');
    case "JPN_JP":
      return LanguageApp.translate(message,'en','ja');
    case "KOR_KR":
      return LanguageApp.translate(message,'en','ko');
    case "POR_BR":
      return LanguageApp.translate(message,'en','pt');
    case "RUS_RU":
      return LanguageApp.translate(message,'en','ru');
    case "SPA_XM":
      return LanguageApp.translate(message,'en','es');
    case "THA_TH":
      return LanguageApp.translate(message,'en','th');
    case "TUR_TR":
      return LanguageApp.translate(message,'en','tr');
    default:
      return message;
  }
}


/**
 * Sort sheet
 * @param {Object} targetSheet - The sheet to sort
 * @param {Array} columns - The columns to sort starting at 0
 */
function sortSheets(targetSheet,columns=[]){
  const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const request = {
      "requests": [
        {
          "sortRange": {
            "range": {
              "sheetId": targetSheet.getSheetId(),
              "startRowIndex": 1
              //"endRowIndex": 10,
              //"startColumnIndex": 0,
              //"endColumnIndex": 6
            },
            "sortSpecs": [
            ]
          }
        }
      ]
  }
  columns.forEach(col => {
    request.requests[0].sortRange.sortSpecs.push({"dimensionIndex": col, "sortOrder": 'ASCENDING'});
  })
  SheetsAPI.Spreadsheets.batchUpdate(request,sheetId);
}

/**
 * Sort sheet
 * @param {Object} targetSheet - The sheet to sort
 * @param {Array} players - List of ally codes to remove
 * @param {Integer} allyCol - The column to find the ally code in
 */
function deleteRows(targetSheet){
  var start = (targetSheet.getLastRow() === 1) ? 2 : targetSheet.getLastRow();
  const request = { 
    "requests": [
      {
        "deleteDimension": {
          "range": {
            "sheetId": targetSheet.getSheetId(),
            "dimension": "ROWS",
            "startIndex": start,
            "endIndex": targetSheet.getMaxRows()
          }
        }
      }
    ]
  };
  if(targetSheet.getLastRow() !== targetSheet.getMaxRows() && start !== targetSheet.getMaxRows()){
    SheetsAPI.Spreadsheets.batchUpdate(request,SpreadsheetApp.getActiveSpreadsheet().getId());
  }
}

/**
 * Clear various rows in a batch
 * @param {Object} targetSheet - The sheet to sort
 * @param {Array} players - List of ally codes to remove
 * @param {Integer} allyCol - The column to find the ally code in
 */
function deleteEntries(targetSheet,players,allyCol = 0,cleanup=false){
  var startTime = new Date();
  //sortSheets(targetSheet,[0]);
  const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var erase = {
    "ranges": []
  };
  var currentData = targetSheet.getRange(2,1,targetSheet.getLastRow(),targetSheet.getLastColumn()).getValues();
  let start = 0;
  let end = 0;
  let row
  let columns = targetSheet.getLastColumn();
  let range
  for(let i=0; i < currentData.length;i++){
    row = i+2;
    if(cleanup){
      if(players[currentData[i][allyCol]] === undefined && currentData[i][allyCol] !==""){
        if(start === 0){
          start = row;
        }
      } else{
        if(start !== 0){
          end = row-1;
          range = targetSheet.getName() + '!R' + start + "C1:R"+ end +"C"+columns;
          erase.ranges.push(range);
          start = 0;
          end = 0;
        }
      }
    } else{
      if(players[currentData[i][allyCol]] !== undefined ){
        if(start === 0){
          start = row;
        }
      } else{
        if(start !== 0){
          end = row-1;
          range = targetSheet.getName() + '!R' + start + "C1:R"+ end +"C"+columns;
          erase.ranges.push(range);
          start = 0;
          end = 0;
        }
      }

    }
  }
  if(end === 0 && start !== 0){
    erase.ranges.push(targetSheet.getName() + '!R' + start + "C1:R"+ targetSheet.getLastRow() +"C"+columns);
  }
  SheetsAPI.Spreadsheets.Values.batchClear(erase,sheetId);
  sortSheets(targetSheet,[0]);
  var endTime = new Date();
  console.log("Deleted data on "+targetSheet.getName()+" in: " + parseInt( (endTime-startTime)/1000));
}


function appendData(targetSheet,data,raw=true){
  var startTime = new Date();
  let inputType = (raw) ? "RAW" : "USER_ENTERED"
  let options = { 'valueInputOption': inputType };
  let values = { 'values': data };
  let range = targetSheet.getName() + '!A1:' + getColumnLetter(targetSheet.getLastColumn()) + '1';
  SheetsAPI.Spreadsheets.Values.append(values, SpreadsheetApp.getActiveSpreadsheet().getId(), range, options);
  var endTime = new Date();
  console.log("Appended "+targetSheet.getName()+" in: " + parseInt( (endTime-startTime)/1000));
}


/**
 * Returns the omicron area based on mode
 * @param {integer} code - The omicron mode
 * @returns {string} The omicron area
 */
function getOmicronArea(mode){
  switch(mode){
    case 0:
      return "";
    case 1: //ALLOMICRON
      return "";
    case 2: //PVEOMICRON
      return "";
    case 3: //PVPOMICRON
      return "";
    case 4:
      return "Raids";
    case 5:
      return "Territory Battles - Combat Mission";
    case 6:
      return "Territory Battles - Special Mission";
    case 7:
      return "Territory Battles";
    case 8:
      return "Territory Wars";
    case 9:
      return "Grand Arena";
    case 10:
      return "Galactic War";
    case 11:
      return "Galactic Conquest";
    case 12:
      return "Galactic Challenges";
    case 13: //PVEEVENTOMICRON
      return "";
    case 14:
      return "Grand Arena - 3v3";
    case 15:
      return "Grand Arena - 5v5";
  }
}

/**
 * 
 */
function getModSearchID(mod){
  var primary = {
    "52":"A", //Accuracy
    "54":"B", //Critical Avoidance
    "53":"C",  //Critical Chance
    "16":"D", //Critical Damage
    "49":"E",  //Defense %
    "55":"F", //Health %
    "48":"G", //Offense %
    "17":"H", //Potency
    "56":"I", //Protection %
    "5":"J", //Speed
    "18":"K" //Tenacity
    };
  var secondary = {
    "53":"L", //Critical Chance
    "42":"M", //Defense
    "49":"N", //Defense %
    "1":"O", //Health
    "55":"P", //Health %
    "41":"Q", //Offense
    "48":"R", //Offense %
    "17":"S", //Potency
    "28":"T", //Protection
    "56":"U", //Protection %
    "5":"V", //Speed
    "18":"W", //Tenacity
  };
  var id = mod[0].toString() + mod[1].toString() + primary[mod[2]];
  if(mod[3] === ""){
    id = id + "0000";
  }else{
    id = id + secondary[mod[3]];
    if(mod[4] === ""){
      id = id + "000";
    }else{
      id = id + secondary[mod[4]];
      if(mod[5] === ""){
        id = id + "00";
      }else{
        id = id + secondary[mod[5]];
        if(mod[6] === ""){
          id = id + "0";
        }else{
          id = id + secondary[mod[6]];
        }
      }
    }
  }
  return id;  
}

/**
 * Returns the column letter for A1 notation
 * @param {Integer} col - The column number
 * @returns {String} The column letter
 */
function getColumnLetter(col){
  switch(col){
    case 1:
      return "A";
    case 2:
      return "B";
    case 3:
      return "C";
    case 4:
      return "D";
    case 5:
      return "E";
    case 6:
      return "F";
    case 7:
      return "G";
    case 8:
      return "H";
    case 9:
      return "I";
    case 10:
      return "J";
    case 11:
      return "K";
    case 12:
      return "L";
    case 13:
      return "M";
    case 14:
      return "N";
    case 15:
      return "O";
    case 16:
      return "P";
    case 17:
      return "Q";
    case 18:
      return "R";
    case 19:
      return "S";
    case 20:
      return "T";
    case 21:
      return "U";
    case 22:
      return "V";
    case 23:
      return "W";
    case 24:
      return "X";
    case 25:
      return "Y";
    case 26:
      return "Z";
    case 27:
      return "AA";
    case 28:
      return "AB";
    case 29:
      return "AC";
    case 30:
      return "AD";
    case 31:
      return "AE";
    case 32:
      return "AF";
    case 33:
      return "AG";
    case 34:
      return "AH";
    case 35:
      return "AI";
    case 36:
      return "AJ";
    case 37:
      return "AK";
    case 38:
      return "AL";
    case 39:
      return "AM";
    case 40:
      return "AN";
    case 41:
      return "AO";
    case 42:
      return "AP";
    case 43:
      return "AQ";
    case 44:
      return "AR";
    case 45:
      return "AS";
    case 46:
      return "AT";
    case 47:
      return "AU";
    case 48:
      return "AV";
    case 49:
      return "AW";
    case 50:
      return "AX";
    case 51:
      return "AY";
    case 52:
      return "AZ";
  }
}
