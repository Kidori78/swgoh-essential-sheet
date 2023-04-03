/*
Functions found here
- removeSelectedGuildData(settingSheetName)
- removeSelectedPlayerData(settingSheetName)
- removeGuildData(settingSheetName,id)
*/

/**
 * Removes all guilds selected on the setting sheet
 */
function removeSelectedGuildData(settingSheetName) {
  //-->Get selected guilds
  const ws = new SheetMap(settingSheetName);
  const guilds = ws.loaded_guilds.filter(g => g[4] !== true);
  const loadedGuilds = ws.loaded_guilds.filter(g => g[4] === true);
  var selectedGuilds = [];
  loadedGuilds.forEach(guild => {
    selectedGuilds[guild[1]] = guild[1];
  });
  const players = ws.loaded_players.filter(player => player[0] !== "").map(player => player[1] );
  var keepPlayers = [];
  players.forEach(player => {
    keepPlayers[ player] = player;
  });
  ws.player_sheet.getRange(2,1,ws.player_sheet.getLastRow(),ws.player_sheet.getLastColumn()).getValues().forEach(player =>{
    if(selectedGuilds[ player[0] ] === undefined && player[0] !== ""){
      keepPlayers[ player[2] ] = player[2];
    }
  });

  if(loadedGuilds.length > 0){
    //-->Remove all data for them
    removeGuildData(settingSheetName,selectedGuilds);
    deleteEntries(ws.player_sheet,keepPlayers,2,true);
    deleteEntries(ws.roster_hero_sheet,keepPlayers,0,true);
    deleteEntries(ws.roster_ship_sheet,keepPlayers,0,true);
    deleteEntries(ws.mod_sheet,keepPlayers,0,true);
    deleteEntries(ws.datacron_sheet,keepPlayers,0,true);
    //-->Delete empty rows
    SpreadsheetApp.flush();
    deleteRows(ws.player_sheet);
    deleteRows(ws.roster_ship_sheet);
    deleteRows(ws.roster_hero_sheet);
    deleteRows(ws.mod_sheet);
    deleteRows(ws.datacron_sheet);

    //-->Remove from settings sheet
    ws.settings_sheet.getRange("B4:F28").clearContent();
    ws.settings_sheet.getRange(4,2,guilds.length, guilds[0].length).setValues(guilds);
  }  
}


/**
 * Removes all players selected on the setting sheet
 */
function removeSelectedPlayerData(settingSheetName){
  //-->Get selected players
  const ws = new SheetMap(settingSheetName);
  const loadedPlayers = ws.loaded_players.filter(p => p[4] === true);
  const players = ws.loaded_players.filter(p => p[4] !== true);
  var removePlayers = [];
  loadedPlayers.forEach(player => {
    removePlayers[player[1]] = player[1];
  });

  if(loadedPlayers.length > 0){
    //-->Remove all data for them
    deleteEntries(ws.player_sheet,removePlayers,2);
    deleteEntries(ws.roster_hero_sheet,removePlayers);
    deleteEntries(ws.roster_ship_sheet,removePlayers);
    deleteEntries(ws.mod_sheet,removePlayers);
    deleteEntries(ws.datacron_sheet,removePlayers);
    //-->Delete empty rows
    SpreadsheetApp.flush();
    deleteRows(ws.player_sheet);
    deleteRows(ws.roster_ship_sheet);
    deleteRows(ws.roster_hero_sheet);
    deleteRows(ws.mod_sheet);
    deleteRows(ws.datacron_sheet);

    //--Remove from settings sheet
    ws.settings_sheet.getRange("B32:F56").clearContent();
    ws.settings_sheet.getRange(32,2,players.length, players[0].length).setValues(players);
  }
}


/**
 * Removes guild from guildData sheet
 * @param {String} id = The guild id
 * @param {String} settingSheetName = The setting sheet name
 */
function removeGuildData(settingSheetName,id){
  const ws = new SheetMap(settingSheetName);
  const targetSheet = ws.guild_sheet;
  const currentData = targetSheet.getRange(2,1,(targetSheet.getLastRow()-1),targetSheet.getLastColumn()).getValues();
  for(let r=0; r < currentData.length; r++){
    if(id[currentData[r][0]] !== undefined){
      targetSheet.getRange(r+2,1,1,targetSheet.getMaxColumns()).clearContent();
    }
  }
  if(targetSheet.getLastRow() !== 1){
    sortSheets(targetSheet,[1]);
    deleteRows(targetSheet);
  }
}
