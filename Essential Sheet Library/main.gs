/*
Functions here:
 - createMenu(language)
 - getPrivacy(language)
 - getDiscord(language)
 - getCellsRemaining(language)
 - buildHelpSheet(language)
 - buildSettingsSheet(language, sheetName)
 */

function createMenu(lang){
  var useLang = getLocalizationLang(lang);
 SpreadsheetApp.getUi()
    .createMenu( translate('Essential Sheet Menu',useLang) )
    .addSubMenu(SpreadsheetApp.getUi().createMenu( translate('Load data',useLang) )
      .addItem( translate('Load guild data',useLang), 'loadGuildData')
      .addItem( translate('Load player data',useLang), 'loadPlayerData')
      .addItem( translate('Load game data',useLang), 'loadGameData')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu( translate('Remove data',useLang) )
      .addItem( translate('Remove guild data',useLang), 'removeSelectedGuildData')
      .addItem( translate('Remove player data',useLang), 'removeSelectedPlayerData')
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu( translate('Help',useLang) )
      .addItem( translate('Calculate remaining cells',useLang), 'getCellsRemaining')
      .addItem( translate('Build Settings sheet',useLang), 'buildSettingsSheet')
      .addItem( translate('Build instructions sheet',useLang), 'buildHelpSheet')
    )
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu( translate('About',useLang) )
      .addItem( translate('Privacy Policy',useLang), 'getPrivacy')
      .addItem( translate('Contact',useLang), 'getDiscord')
    )
    .addToUi();
} 

/**
 * Displays the privacy policy
 * @param {string} language - The language to show the message in.
 */
function getPrivacy(language){
  SpreadsheetApp.getUi().alert(translate("PRIVACY POLICY INFORMATION:\n\nThe Essential Sheet was created by Kidori and does not include any programming code, add-ons, or any other means for the creator of it to collect any data of any kind from any users who choose to use it. The code for the library can be reviewed on GitHub at https://github.com/Kidori78/swgoh-essential-sheet.", language));
}

/**
 * Displays the privacy policy
 * @param {string} language - The language to show the message in.
 */
function getDiscord(language){
  SpreadsheetApp.getUi().alert(translate("CONTACT INFORMATION:\n\nThe Essential Sheet was created by Kidori and you can find additional help and support for it on ", language) + "Kidori's SWGoH Tools Discord: https://discord.gg/Z3fWRWa77W");
}

/**
 * Gets the remaining number of cells the sheet can have.
 * @param {string} language - The language to show the message in.
 */
function getCellsRemaining(language){
  var useLang = getLocalizationLang(language);
  var cellCount = 0;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sheet => {
    cellCount += sheet.getMaxColumns() * sheet.getMaxRows();
  })
  SpreadsheetApp.getUi().alert(translate("There are currently " +cellCount+" cells in this Google Sheet. There are " + (10000000 - cellCount) + " cells remaining.",useLang));
}


/******************************************************
 * Builds the instruction sheet.
 * @param {string} lang - The language to set the sheet to
 */
function buildHelpSheet(lang){
  const ws = SpreadsheetApp.getActiveSpreadsheet();
  var useLang = getLocalizationLang(lang);
  const shtName = translate("Essential Instructions",useLang);
  if(ws.getSheetByName(shtName) === null){
    //Create Sheet
    ws.insertSheet(shtName);
  } 
  else{ //Clear sheet
    ws.getSheetByName(shtName).clearContents().clearFormats().clear().getRange("A1:K59").clearDataValidations();
  }
  const instructionSheet = ws.getSheetByName(shtName);
  var writeData = [];

  writeData.push(
    ["SWGoH Essential Sheet",""], 
    [translate("Created by Kidori",useLang), ""],
    ["",""],
    [translate("This sheet will allow you to utilize Comlink to get information straight from Star Wars Galaxy of Heroes and store it to use with your own tools or projects. It will require a desktop to load the data.",useLang), ""],
    [translate("How to use",useLang),""],
    [translate("Step 1",useLang), translate("To get started go to File then Make a copy in the menu bar."+ String.fromCharCode(10) +"If your preferred language is English then skip to Step 5. In your new copy go to the menu bar and select Extensions then Apps Script",useLang) ],
    [translate("Step 2",useLang), translate("In Code.gs adjust the sheetName within the PropertiesService to whatever supported language you want the sheet to use."+String.fromCharCode(10)+"Default is essentialSettings.",useLang)],
    [translate("Step 3",useLang),translate("In Code.gs adjust the language within the PropertiesServce to whatever supported language you want the sheet to use."+String.fromCharCode(10)+"Default is English.",useLang)],
    [translate("Step 4",useLang), translate("Refresh the browser.",useLang)],
    [translate("Step 5",useLang), translate("In the menu bar select Essential Sheet Menu => Help, then Build Settings Sheet. It will require you to authorize the functionalities needed to use this sheet."+String.fromCharCode(10)+ "On the pop-up screen choose your account and on the next screen you will need to click Advanced at the bottom. It will ask if you are sure and you must click the link at the very bottom that says Go to SWGoH Essential Sheet (unsafe), it may have your custom sheets name. Under all of the authorization details you will need to click Allow. If you are concerned about safety there is a detailed FAQ about authorization in Kidori's SWGoH Tools discord."+String.fromCharCode(10)+String.fromCharCode(10)+"After accepting the authorization go back to the menu bar and select Essential Sheet Menu then Build Settings Sheet again.",useLang)],
    [translate("Step 6",useLang),translate("Fill out all api settings and sheet options on the Essential Settings sheet.",useLang)],
    [translate("Step 7",useLang),translate("Go to Essential Sheet Menu then Load data then Load game data.",useLang)],
    [translate("Step 8",useLang), translate("After that enter the guild id or player id on the settings sheet and then go to the menu bar and select Essential Sheet Menu then Load data then Load guild data or Load player data depending on which one you want to get.",useLang)],
    ["",""],
    [translate("If you just want to insert the Essential Sheet into an existing project, copy the Code.gs file in this sheets Apps Script and paste it into your existing sheet. You will need to add Google Sheets API to the Services by clicking the + next to it in the sidebar within Apps Script. You will select v4 and then rename the identifier to SheetsAPI. Then add the following Library Script ID to the project:",useLang),""],
    ["","1Ygs0Z2bkJLSt9in4QGaDFNEq1VzhAfKuumZClQuWY45c-9Kx7je87kih"],
    ["",""],
    [translate("Advanced Uses",useLang), ""],
    [translate("Adjust App Script by adding functions that utilize available libraries and functions. The Essential Sheet library has the following libraries:",useLang),""],
    ["SWGoH Essential Sheet",""],
    [translate("Library Id:",useLang),"EssentialSheet " + translate("Note: You are able to change this to whatever you want.", useLang)],
    [translate("Documentation",useLang), "https://github.com/Kidori78/swgoh-essential-sheet"],
    ["SWGoH Comlink API",""],
    [translate("Library ID:",useLang), "SWGoHAPI"],
    [translate("Documentation:",useLang),"https://github.com/Kidori78/swgoh-comlink-wrapper-for-gas"],
    ["SWGoH Stat Calculator",""],
    [translate("Library ID:",useLang), "StatCalculator"],    
    [translate("Documentation:",useLang),"https://github.com/Kidori78/swgoh-stat-calc/tree/GAS-v2"],
    ["Google Sheets API",""],
    [translate("Library ID:",useLang), "SheetsAPI"],    
    [translate("Documentation:",useLang),"https://developers.google.com/sheets/api/guides/concepts"],
    ["",""],
    [translate("To utilize these services you will need to grab them from the Essential Library by calling it first:",useLang),""],
    ["essentialLibrary.installedLibrary.function/class",""],
    [translate("Examples",useLang),""],
    [translate("Class",useLang), "var api = new EssentialSheet.SWGoHAPI.Comlink(arguments)"],
    [translate("Class",useLang), "var statCalc = new EssentialSheet.StatCalculator.StatCalculator(arguments)"],
    [translate("Function",useLang), "EssentialSheet.loadGameData(arguments)"],
    [translate("Function",useLang), "EssentialSheet.SWGoHAPI.comlinkAPI_help()"],
    ["",""],
    [translate("Use Case",useLang),""],
    [translate("You can extend any of the provided data sheets or to build additional sheets.",useLang), ""],
    [translate("Example",useLang),""],
    [translate("You can create a new menu at the top with your own script to execute that can add a column to unitData for unit locations. Then when you need to update the unitData sheet you can call it directly passing special arguments to it so it can still update units without deleting the new location column.",useLang),""],
    ["",""],
    [translate("All functions for creating the game data sheets have a parameter for appending added columns to the data as well as returning the game data collections to minimize fetches.",useLang),""],
    ["addUnitLocations() { ... }",""],
    ["updateUnitDataSheet(settingSheetName,gameData,append) => updateDataSheet(settingSheetName, null, true",""]
  );

  instructionSheet.getRange(1,1,writeData.length, writeData[0].length).setValues(writeData);
  //Format Sheet
  instructionSheet.getRangeList(["A1","A2","A5:A13","B16","A18","A20","A23", "A26", "A29", "A35", "A41", "A43"]).setFontWeight('bold');
  instructionSheet.getRange("A1").setFontSize(16);
  instructionSheet.getRange("A5:J17").setBorder(true,false,true,false,false,false,'black',SpreadsheetApp.BorderStyle.SOLID);
  let range = ["A4:H4","B6:I6","B7:I7","B8:I8","B9:I9","B10:I10","B11:I11","B12:I12","B13:I13","A15:I15","A44:H44","A46:H46"];
  range.forEach(rng => {
    instructionSheet.getRange(rng).mergeAcross().setWrap(true);
  })
  instructionSheet.getRange("A6:I13").setVerticalAlignment('middle').setBorder(true,null,true,null,null,true,'#CCCCCC',SpreadsheetApp.BorderStyle.DASHED);
  deleteUnusedCells(instructionSheet,7); 
  //instructionSheet.deleteColumns(9,17);
  //instructionSheet.deleteRows(50, 950);
  instructionSheet.setHiddenGridlines(true);
}

/******************************************************
 * Builds the settings sheet.
 * @param {string} lang - The language to set the sheet to
 * @param {string} shtName - The name for the sheet
 */
function buildSettingsSheet(lang,shtName=null){
  const ws = SpreadsheetApp.getActiveSpreadsheet();
  var useLang = getLocalizationLang(lang);
  if(shtName === null) { shtName = "essentialSettings"; }
  if(ws.getSheetByName(shtName) === null){
    //Create Sheet
    ws.insertSheet(shtName);
  } 
  else{ //Clear sheet
    ws.getSheetByName(shtName).clearContents().clearFormats().clear().getRange("A1:K60").clearDataValidations();
  }
  const settingSheet = ws.getSheetByName(shtName);

  //Add data
  settingSheet.getRange("B2:F3").setValues([[translate("LOADED GUILDS",useLang),"","","",""],[translate("Guild Name",useLang), translate("Guild ID",useLang), translate("Notes",useLang), translate("Last Sync",useLang), translate("Remove",useLang)]]);
  settingSheet.getRange("F4:F28").insertCheckboxes();
  settingSheet.getRange("B30:F31").setValues([[translate("LOADED PLAYERS",useLang),"","","",""],[translate("Player Name",useLang), translate("Ally Code",useLang), translate("Notes",useLang), translate("Last Sync",useLang), translate("Remove",useLang)]]);
  settingSheet.getRange("F32:F56").insertCheckboxes();
  settingSheet.getRange("H2:I59").setValues([
    [translate("INSTRUCTIONS",useLang),""],
    [translate("Select API settings and preferred options",useLang),""],
    [translate('In Get Guild/Player enter the appropriate id.',useLang),""],
    [translate('In the menu bar select Essential Sheet Menu',useLang),""],
    [translate("Select Load data, choose data to load.",useLang),""],
    ["",""],
    [translate("",useLang), translate("Ally Code | Player ID | Guild ID",useLang)],
    [translate("GET GUILD",useLang), ""],
    [translate("GET PLAYER",useLang),""],
    ["",""],
    ["Comlink API "+translate("Setup",useLang),""],
    [translate("Host URL",useLang),""],
    [translate("Access Key",useLang),""],
    [translate("Secret Key",useLang),""],
    [translate("Language",useLang),""],
    ["",""],
    [translate("Available Sheets",useLang),""],
    [translate("Description",useLang),translate("Include",useLang)],
    [translate("Guild Data",useLang),""],
    [translate("Player/Member Data",useLang),""],
    [translate("Roster Hero Data",useLang),""],
    [translate("Roster Ship Data",useLang),""],
    [translate("Mod Data",useLang),""],
    [translate("Datacron Data",useLang),""],
    [translate("Unit Data",useLang),""],
    [translate("Hero Base Data",useLang),""],
    [translate("Hero Ability Data",useLang),""],
    [translate("Ship Ability Data",useLang),""],
    [translate("Gear Data",useLang),""],
    [translate("Relic Data",useLang),""],
    ["",""],
    [translate("SHEET OPTIONS",useLang),""],
    [translate("Guild Data Details",useLang),""],
    [translate("Event Activity",useLang),""],
    [translate("Member Data Details",useLang),""],
    [translate("Galactic Power",useLang),""],
    [translate("Times",useLang),""],
    [translate("Grand Arena",useLang),""],
    [translate("Raid Tickets",useLang),""],
    [translate("Raid Activity",useLang),""],
    [translate("Roster Character Data Details",useLang),""],
    [translate("Stats",useLang),""],
    [translate("Zetas/Omicrons",useLang),""],
    [translate("Ability Tiers",useLang),""],
    [translate("Mod Data Details",useLang),""],
    [translate("Raw Roll Values",useLang),""],
    [translate("Stats",useLang),""],
    [translate("Roster Ship Data Details",useLang),""],
    [translate("Stats",useLang),""],
    [translate("Ability Tiers",useLang),""],
    [translate("Unit Data Details",useLang),""],
    [translate("Images",useLang),""],
    [translate("Unit Base Data Details",useLang),""],
    [translate("Stats",useLang),""],
    [translate("Ability Data Details",useLang),""],
    [translate("Images",useLang),""],
    [translate("Tier Details",useLang),""],
    [translate("Images provided by https://swgoh.gg",useLang),""]
  ]);
  settingSheet.getRange("J19:J31").setValues([
    [translate("Name",useLang)],
    ["guildData"],
    ["playerData"],
    ["rosterHeroData"],
    ["rosterShipData"],
    ["modData"],
    ["datacronData"],
    ["unitData"],
    ["heroBaseData"],
    ["heroAbilityData"],
    ["shipAbilityData"],
    ["gearData"],
    ["relicData"]
  ]);
  settingSheet.getRangeList(["I22:I31","I35","I37:I41","I43:I45","I47","I48","I50","I51","I53","I55","I57","I58"]).insertCheckboxes();
  var rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList([
        "English",
        "中文(简体)",
        "中國傳統的",
        "Français",
        "Deutsch",
        "Indonesia",
        "Italiano",
        "日本語",
        "한국인",
        "Português (Brasil)",
        "Русский",
        "Español",
        "แบบไทย",
        "Türkçe"])
    .setAllowInvalid(false)
    .build();
  settingSheet.getRange('I16').setDataValidation(rule);
  settingSheet.getRange('I16').setValue(lang);

  //Format Sheet
  settingSheet.getRange("A1:K60").setBackground('#666666');
  //Widths
  settingSheet.setColumnWidth(1, 30);
  settingSheet.setColumnWidths(2, 4, 150);
  settingSheet.setColumnWidth(6, 70);
  settingSheet.setColumnWidth(7, 50);
  settingSheet.setColumnWidth(8, 120);
  settingSheet.setColumnWidth(9, 70);
  settingSheet.setColumnWidth(10, 130);
  settingSheet.setColumnWidth(11, 30);
  //Colors
  settingSheet.getRange("H59:J59").merge().setFontColor('#ffe599').setFontWeight('bold').setHorizontalAlignment('center');
  settingSheet.getRangeList(["I9:J10","I13:J16"]).setBackground('#ffffff');
  settingSheet.getRange("I20:J31").setBackground('#ffffff').setFontColor('#e69138').setFontWeight('bold');
  settingSheet.getRange("I33:J58").setBackground('#ffffff').setFontColor('#45818e');
  settingSheet.getRangeList(["F3","F31"]).setBackground('#e06666').setFontWeight('bold').setHorizontalAlignment('center');
  settingSheet.getRangeList(["F4:F28","F32:F56"]).setBackground('#f4cccc');
  settingSheet.getRange("H12:H16").setBackground("#c27ba0").setFontWeight('bold').setFontColor("#ead1dc");
  settingSheet.getRange("H19:H31").setBackground("#f6b26b").setFontWeight('bold').setFontColor("#434343");
  settingSheet.getRange("H26:H31").setBackground('#f9cb9c');
  settingSheet.getRange("H33:H58").setBackground("#a2c4c9").setFontWeight('bold').setFontColor("#434343");
  let mergeRanges = ["B2:F2", "H2:J2", "H8", "I8:J8", "H12:J12", "H18:J18", "H33:J33", "H34:J34", "H36:J36", "H42:J42", "H46:J46", "H49:J49", "H52:J52", "H54:J54", "H56:J56", "B30:F30"];
  let headerColors = ["#6aa84f","#b7b7b7","#93c47d","#b6d7a8","#a64d79","#e69138", "#45818e", "#76a5af", "#76a5af", "#76a5af", "#76a5af", "#76a5af", "#76a5af", "#76a5af", "#76a5af", "#3d85c6"];
  let headerFontColor = ["#ffffff","#000000","#000000","#000000","#ead1dc","#000000","#d0e0e3","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#cfe2f3"];
  for(let i=0; i < mergeRanges.length;i++){
    settingSheet.getRange(mergeRanges[i]).merge().setBackground(headerColors[i]).setFontColor(headerFontColor[i]);
  }
  settingSheet.getRange("H19:J19").setBackground('#e69138').setFontWeight('bold').setFontColor('black').setHorizontalAlignment('center');
  settingSheet.getRange("H9:H10").setBackground('#93c47d').setFontWeight('bold');
  settingSheet.getRange("I9:J9").merge();
  settingSheet.getRange("I10:J10").merge();
  settingSheet.getRange("I13:J16").mergeAcross();
  settingSheet.getRangeList(mergeRanges).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  settingSheet.getRange("B3:E28").applyRowBanding().setHeaderRowColor('#8bc34a').setFirstRowColor('#d9ead3').setSecondRowColor('#eef7e3');
  settingSheet.getRange("B3:E3").setFontWeight('bold').setHorizontalAlignment('center');
  settingSheet.getRange("B31:E56").applyRowBanding().setHeaderRowColor('#6fa8dc').setFirstRowColor('#9fc5e8').setSecondRowColor('#cfe2f3');
  settingSheet.getRange("B31:E31").setFontWeight('bold').setHorizontalAlignment('center');
  //Borders
  settingSheet.getRangeList(["B2:F28", "B30:F56", "H2:J6", "H8:J10", "H12:J16", "H18:J31", "H33:J58"]).setBorder(true,true,true,true,null,null,'black',SpreadsheetApp.BorderStyle.SOLID_THICK).setBorder(null,null,null,null,null,true,'black',SpreadsheetApp.BorderStyle.SOLID);
  settingSheet.getRangeList(["B2:F28", "B30:F56", "H8:I10", "H19:J31"]).setBorder(null,null,null,null,true,null,'black',SpreadsheetApp.BorderStyle.SOLID);
  settingSheet.getRange("H3:J6").setBorder(null,null,null,null,null,false).setBackground('#d9d9d9');
  settingSheet.getRange("H26:J26").setBorder(true,null,null,null,null,null,'black',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  settingSheet.setHiddenGridlines(true);
  settingSheet.getRange("H20:I20").merge().setHorizontalAlignment('center');
  settingSheet.getRange("H21:I21").merge().setHorizontalAlignment('center');
  settingSheet.getRange("K20:K25").merge().setVerticalAlignment('middle').setHorizontalAlignment('center').setValue(translate("User Data",useLang)).setTextRotation(90).setFontColor('#ffd966').setFontWeight('bold');
  settingSheet.getRange("K26:K31").merge().setVerticalAlignment('middle').setHorizontalAlignment('center').setValue(translate("Game Data",useLang)).setTextRotation(90).setFontColor('#ffd966').setFontWeight('bold');
  //Delete extra rows and columns
  deleteUnusedCells(settingSheet);
}
