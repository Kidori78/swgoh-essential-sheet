/**
 * To localize the menu select the language from the following options and then change the language of PropertiesService:
 * English, 中文(简体), 中國傳統的, Français, Deutsch, Indonesia, Italiano, 日本語, 한국인, Português (Brasil), Русский, Español, แบบไทย, Türkçe
 * To change the settings sheet name adjust the sheetName below.
 */
PropertiesService.getUserProperties().setProperties({sheetName: "essentialSettings", language: "English"});

function onOpen(){
  var language = PropertiesService.getUserProperties().getProperty('language');
  EssentialSheet.createMenu(language);
}

function buildSettingsSheet(){
  var settingSheetName = PropertiesService.getUserProperties().getProperty('sheetName');
  var language = PropertiesService.getUserProperties().getProperty('language');
  EssentialSheet.buildSettingsSheet(language,settingSheetName);
}

function buildHelpSheet(){
  var language = PropertiesService.getUserProperties().getProperty('language');
  EssentialSheet.buildHelpSheet(language);
}

function loadGuildData(){
  var settingSheetName = PropertiesService.getUserProperties().getProperty('sheetName');  
  EssentialSheet.loadGuildData(settingSheetName);
}

function loadPlayerData(){
  var settingSheetName = PropertiesService.getUserProperties().getProperty('sheetName');  
  EssentialSheet.loadPlayerData(settingSheetName);
}

function loadGameData(){
  var settingSheetName = PropertiesService.getUserProperties().getProperty('sheetName');
  EssentialSheet.loadGameData(settingSheetName);
}

function removeSelectedGuildData(){
  var settingSheetName = PropertiesService.getUserProperties().getProperty('sheetName');
  EssentialSheet.removeSelectedGuildData(settingSheetName);
}

function removeSelectedPlayerData(){
  var settingSheetName = PropertiesService.getUserProperties().getProperty('sheetName');
  EssentialSheet.removeSelectedPlayerData(settingSheetName);
}

function getCellsRemaining(){
  var language = PropertiesService.getUserProperties().getProperty('language');
  EssentialSheet.getCellsRemaining(language);
}

function getPrivacy(){
  var language = PropertiesService.getUserProperties().getProperty('language');
  EssentialSheet.getPrivacy(language);
}

function getDiscord(){
  var language = PropertiesService.getUserProperties().getProperty('language');
  EssentialSheet.getDiscord(language);
}
