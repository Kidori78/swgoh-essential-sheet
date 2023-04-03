/*
Functions found here
- loadGuildData(settingSheetName)
- loadPlayerData(settingSheetName)
- updateGuildDataSheet(settingSheetName, gameData)
- updatePlayerDataSheet(settingSheetName, gameData)
- updateHeroRosterDataSheet(settingSheetName, gameData)
- updateShipRosterDataSheet(settingSheetName, gameData)
- updateModDataSheet(settingSheetName, gameData)
- updateDatacronDataSheet(settingSheetName, gameData)
*/

/**
 * Builds the guild data sheets with the selected options from the essentialSettings tab.
 * @param {string} settingSheetName - The name of the settings sheet
 */
function loadGuildData(settingSheetName){
  //Verify sheets to create
  const ws = new SheetMap(settingSheetName);
  const create = {
    "guildData": { "include": true, "sheetName": ws.name_guildSheet },
    "rosterHeroData": { "include": ws.create_heroRoster, "sheetName": ws.name_rosterHeroSheet },
    "rosterShipData": { "include": ws.create_shipRoster, "sheetName": ws.name_rosterShipSheet },
    "modData": { "include": ws.create_mod, "sheetName": ws.name_modSheet },
    "datacronData": { "include": ws.create_datacron, "sheetName": ws.name_datacronSheet }
  };
  var guildData = null;
  var globalStart = new Date();
  var segmentStart
  var segmentEnd
  //Check if sheets exist
  for(let key in create){
    if(create[key].include){
      //Build Sheet
      switch(key){
        case "guildData":
          segmentStart = new Date();
          guildData = updateGuildDataSheet(settingSheetName, guildData);
          segmentEnd = new Date();
          console.log("Guild data loaded in: " + parseInt( (segmentEnd-segmentStart)/1000));
          break;
        case "rosterHeroData":
          segmentStart = new Date();
          guildData = updateHeroRosterDataSheet(settingSheetName, guildData);
          segmentEnd = new Date();
          console.log("Hero Roster data loaded in: " + parseInt( (segmentEnd-segmentStart)/1000));
          break;
        case "rosterShipData":
          segmentStart = new Date();
          guildData = updateShipRosterDataSheet(settingSheetName, guildData);
          segmentEnd = new Date();
          console.log("Ship Roster data loaded in: " + parseInt( (segmentEnd-segmentStart)/1000));
          break;
        case "modData":
          segmentStart = new Date();
          guildData = updateModDataSheet(settingSheetName, guildData);
          segmentEnd = new Date();
          console.log("Mod data loaded in: " + parseInt( (segmentEnd-segmentStart)/1000));
          break;
        case "datacronData":
          segmentStart = new Date();
          guildData = updateDatacronDataSheet(settingSheetName, guildData);
          segmentEnd = new Date();
          console.log("Datacron data loaded in: " + parseInt( (segmentEnd-segmentStart)/1000));
          break;
      }
    }
    var globalEnd = new Date();
          console.log("Global runtime: " + parseInt( (globalEnd-globalStart)/1000));
  }
  //Write loaded guild to settings
  var loadedGuilds = ws.loaded_guilds;
  var loadedMap = [];
  var currentLoaded = [];
  for(let i=0; i < loadedGuilds.length;i++){
    if(loadedGuilds[i][0] !== ""){
      loadedMap[loadedGuilds[i][1]] = i;
      currentLoaded.push(loadedGuilds[i]);
    }
  }
  guildData.forEach(guild => {
    if(loadedMap[guild.profile.id] > -1){
      currentLoaded[loadedMap[guild.profile.id]] = [guild.profile.name,guild.profile.id,loadedGuilds[loadedMap[guild.profile.id]][2],new Date(),false ];
    }else {
      currentLoaded.push([guild.profile.name,guild.profile.id,"",new Date(),false]);
    }
  });
  currentLoaded.sort((a,b) => {
    return a[1] > b[1];
  });
  ws.settings_sheet.getRange(4, 2, currentLoaded.length, currentLoaded[0].length).setValues(currentLoaded);  
}


/**
 * Builds the player data sheets with the selected options from the essentialSettings tab.
 * @param {string} settingSheetName - The name of the settings sheet
 */
function loadPlayerData(settingSheetName){
  //Verify sheets to create
  const ws = new SheetMap(settingSheetName);
  const create = {
    "playerData": { "include": true, "sheetName": ws.name_playerSheet },
    "rosterHeroData": { "include": ws.create_heroRoster, "sheetName": ws.name_rosterHeroSheet },
    "rosterShipData": { "include": ws.create_shipRoster, "sheetName": ws.name_rosterShipSheet },
    "modData": { "include": ws.create_mod, "sheetName": ws.name_modSheet },
    "datacronData": { "include": ws.create_datacron, "sheetName": ws.name_datacronSheet }
  };
  var globalStart = new Date();
  var playerData = null;
  //Check if sheets exist
  for(let key in create){
    if(create[key].include){
      //Build Sheet
      switch(key){
        case "playerData":
          playerData = updatePlayerDataSheet(settingSheetName, playerData);
          break;
        case "rosterHeroData":
          playerData = updateHeroRosterDataSheet(settingSheetName, playerData);
          break;
        case "rosterShipData":
          playerData = updateShipRosterDataSheet(settingSheetName, playerData);
          break;
        case "modData":
          playerData = updateModDataSheet(settingSheetName, playerData);
          break;
        case "datacronData":
          playerData = updateDatacronDataSheet(settingSheetName, playerData);
          break;
      }
    }
    var globalEnd = new Date();
    console.log("Global runtime: " + parseInt( (globalEnd-globalStart)/1000));
  }

  //Write loaded player(s) to settings
  var loadedPlayers = ws.loaded_players;
  var loadedMap = [];
  var currentLoaded = [];
  for(let i=0; i < loadedPlayers.length;i++){
    if(loadedPlayers[i][0] !== ""){
      loadedMap[loadedPlayers[i][1]] = i;
      currentLoaded.push(loadedPlayers[i]);
    }
  }
  playerData.forEach(player => {
    if(loadedMap[player.allyCode] > -1){
      currentLoaded[loadedMap[player.allyCode]] = [player.name,player.allyCode,loadedPlayers[loadedMap[player.allyCode]][2],new Date(),false ];
    }else {
      currentLoaded.push([player.name,player.allyCode,"",new Date(),false]);
    }
  });
  currentLoaded.sort((a,b) => {
    return a[1] > b[1];
  });
  ws.settings_sheet.getRange(32, 2, currentLoaded.length, currentLoaded[0].length).setValues(currentLoaded);
}


/**
 * Creates or adds player data to the playerData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} playerData - The currently retrieved player data
 * @returns {object} playerData - Returns the player data to be used with other functions.
 */
function updatePlayerDataSheet(settingSheetName, playerData = null){
  const ws = new SheetMap(settingSheetName);
  const options = { 
    "gp": ws.option_member_gp,
    "times": ws.option_member_time,
    "gac": ws.option_member_gac,
    "tickets": ws.option_member_tickets,
    "raids": ws.option_member_raids
  };
  //-->Create Sheet and designate write sheet
  var writeSheet
  if(ws.player_sheet === null){
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_playerSheet);
    SpreadsheetApp.flush();
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_playerSheet);
    writeSheet.setFrozenRows(1);
  } else {
    writeSheet = ws.player_sheet;
  }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  var isGuild = false;
  var ids = [];
  if(playerData === null){
    if(ws.input_playerID.toString().indexOf(",") > -1){
      ids = ws.input_playerID.toString().split(",");
    } else{
      ids = [ws.input_playerID.toString()];
    }
    playerData = api.fetchPlayers(ids,false,true);
  } else{
    if(playerData[0].hasOwnProperty("inviteStatus")){
      isGuild = true;
    }
  }
  //-->Build headers
  let headers = [
      translate("guild id",ws.api_language),
      translate("guild name",ws.api_language),
      translate("ally code",ws.api_language),
      translate("name",ws.api_language),
      translate("level",ws.api_language)
  ];
  if(options.gp){
    headers.push(
      translate("galactic power",ws.api_language),
      translate("hero gp",ws.api_language),
      translate("ship gp",ws.api_language)
    );
  }
  if(options.times){
    headers.push(
      translate("last login",ws.api_language),
      translate("guild join time",ws.api_language)
    );
  }
  if(options.gac){
    headers.push(
      translate("gac league",ws.api_language),
      translate("gac division",ws.api_language),
      translate("gac skill rating",ws.api_language),
      translate("gac division rank",ws.api_language)
    );
  }
  if(options.tickets){
    headers.push(
      translate("lifetime raid tickets",ws.api_language),
      translate("current raid tickets",ws.api_language)
    );
  }
  if(options.raids){
    headers.push(
      translate("raid activity",ws.api_language)
    );
  }
  writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).clearContent();
  writeSheet.getRange(1,1,1, headers.length).setValues([headers]);
  writeSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');

  //-->Build Player Data
  var writeData = [];
  var tempPlayer = [];
  var removeList = [];
  if(isGuild){
    if(writeSheet.getRange("A2").getValue() !== ""){
      writeSheet.getRange(2,1,writeSheet.getLastRow() + 1, writeSheet.getLastColumn()).getValues().forEach(player => {
        if(playerData[0].profile.id === player[0]){
          removeList[player[2]] = player[2];
        }
      });
    }
    playerData[0].member.forEach(player => {
      removeList[player.allyCode] = player.allyCode;
      tempPlayer = [];
      tempPlayer.push(
        playerData[0].profile.id, 
        playerData[0].profile.name, 
        player.allyCode,
        player.playerName,
        player.playerLevel
      );
      if(options.gp){
        tempPlayer.push(player.profileStat.filter(stat => stat.nameKey === "Galactic Power:" || stat.nameKey === "STAT_GALACTIC_POWER_ACQUIRED_NAME")[0].value);
        tempPlayer.push(player.profileStat.filter(stat => stat.nameKey ===  "Galactic Power (Characters):" || stat.nameKey === "STAT_CHARACTER_GALACTIC_POWER_ACQUIRED_NAME")[0].value);
        tempPlayer.push(player.profileStat.filter(stat => stat.nameKey === "Galactic Power (Ships):" || stat.nameKey === "STAT_SHIP_GALACTIC_POWER_ACQUIRED_NAME")[0].value);
      }
      if(options.times){
        tempPlayer.push('= ( ('+player.lastActivityTime+'/1000)/86400 ) + ( DATEVALUE("1-1-1970") - DATEVALUE("12-30-1899") )');
        tempPlayer.push('= ( ('+(player.guildJoinTime *1000)+'/1000)/86400 ) + ( DATEVALUE("1-1-1970") - DATEVALUE("12-30-1899") )');
      }
      if(options.gac){
        tempPlayer.push((player.playerRating.playerRankStatus === null) ? "" : player.playerRating.playerRankStatus.leagueId);
        tempPlayer.push((player.playerRating.playerRankStatus === null) ? "" : player.playerRating.playerRankStatus.divisionId);
        tempPlayer.push((player.playerRating.playerSkillRating === null) ? "" : player.playerRating.playerSkillRating.skillRating);
        tempPlayer.push((player.seasonStatus[0] === undefined) ? "" : player.seasonStatus[(player.seasonStatus.length - 1)].rank);
      }
      if(options.tickets){
        tempPlayer.push(player.memberContribution.filter(contrib => contrib.type === 2)[0].lifetimeValue);
        tempPlayer.push(player.memberContribution.filter(contrib => contrib.type === 2)[0].currentValue);
      }
      if(options.raids){
        let tempRaid = "";
        playerData[0].recentRaidResult.forEach(raid => {
          let raidDamage = raid.raidMember.filter(member => member.playerId === player.playerId)[0];
          if(raidDamage === undefined){ 
            raidDamage = 0; 
          } else{ 
            raidDamage = raidDamage.memberProgress; 
          }
          if(raid.outcome === 5){
            raidDamage = "SIM";
          }
          tempRaid += raid.raidId + ": " + raidDamage + ", ";
        });
        tempRaid = tempRaid.substring(0, tempRaid.length -2);
        tempPlayer.push(tempRaid);
      }
        writeData.push(tempPlayer);
    });
  }else{
    playerData.forEach(player => {
      removeList[player.allyCode] = player.allyCode;
      tempPlayer = [];
      tempPlayer.push(
        player.guildId,
        player.guildName,
        player.allyCode,
        player.name,
        player.level
      );
      if(options.gp){
        tempPlayer.push(player.profileStat.filter(stat => { return stat.nameKey === "Galactic Power:" || stat.nameKey === "STAT_GALACTIC_POWER_ACQUIRED_NAME" })[0].value);
        tempPlayer.push(player.profileStat.filter(stat => { return stat.nameKey ===  "Galactic Power (Characters):" || stat.nameKey === "STAT_CHARACTER_GALACTIC_POWER_ACQUIRED_NAME" })[0].value);
        tempPlayer.push(player.profileStat.filter(stat => { return stat.nameKey === "Galactic Power (Ships):" || stat.nameKey === "STAT_SHIP_GALACTIC_POWER_ACQUIRED_NAME"})[0].value);
      }
      if(options.times){
        tempPlayer.push('= ( ('+player.lastActivityTime+'/1000)/86400 ) + ( DATEVALUE("1-1-1970") - DATEVALUE("12-30-1899") )');
        tempPlayer.push(""); //only in guild endpoint
      }
      if(options.gac){
        tempPlayer.push((player.playerRating.playerRankStatus === null) ? "" : player.playerRating.playerRankStatus.leagueId);
        tempPlayer.push((player.playerRating.playerRankStatus === null) ? "" : player.playerRating.playerRankStatus.divisionId);
        tempPlayer.push((player.playerRating.playerSkillRating === null) ? "" : player.playerRating.playerSkillRating.skillRating);
        tempPlayer.push((player.playerRating.playerRankStatus === null) ? "" : player.seasonStatus[(player.seasonStatus.length-1)].rank);
      }
      if(options.tickets){
        tempPlayer.push(""); //only in guild endpoint
        tempPlayer.push(""); //only in guild endpoint
      }
      if(options.raids){
        tempPlayer.push(""); //only in guild endpoint
      }
      writeData.push(tempPlayer);
    });

  }

  //-->Delete old data
  deleteEntries(writeSheet,removeList,2);
  //-->Write player data to sheet
  if(writeData.length > 0){
    appendData(writeSheet,writeData,false);
  }
  //-->Format Sheet
  if(options.times){
    writeSheet.getRangeList(["I2:I","J2:J"]).setNumberFormat('yyyy/mm/dd hh:mm');
  }
  //sortSheets(writeSheet,[1,3]);
  deleteUnusedCells(writeSheet);

  return playerData
}


/**
 * Creates or adds guild data to the guildData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} guildData - The currently retrieved guild data
 * @returns {object} playerData - Returns the player data to be used with other functions.
 */
function updateGuildDataSheet(settingSheetName, guildData = null){
  const ws = new SheetMap(settingSheetName);
  const options = {
    "events": ws.option_guild_event
  };
  //-->Create Sheet and designate write sheet
  var writeSheet
  var currentData = [];
  var currentMap = [];
  if(ws.guild_sheet === null){
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_guildSheet);
    SpreadsheetApp.flush();
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_guildSheet);
    writeSheet.setFrozenRows(1);
  } else {
    writeSheet = ws.guild_sheet;
    currentData = (writeSheet.getRange("A1") === "") ? [] : writeSheet.getRange(1,1,writeSheet.getLastRow(),writeSheet.getLastColumn()).getValues();
    for(var i=1;i < currentData.length;i++){
      currentMap[ currentData[i][0]] = i + 1; //guildId
    }
  }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  var ids = [];
  if(guildData === null){
    ids = [ws.input_guildID.toString()];
    guildData = api.fetchGuildRosters(ids,false,true);
  }
  //-->Build headers
  let headers = [
      translate("guild id",ws.api_language),
      translate("guild name",ws.api_language),
      translate("leader",ws.api_language),
      translate("leader id",ws.api_language),
      translate("galactic power",ws.api_language),
      translate("members",ws.api_language),
      translate("reset time",ws.api_language)
  ];
  if(options.events){
    headers.push(
      translate("raids completed",ws.api_language),
      translate("territory battles",ws.api_language),
      translate("territory war 1",ws.api_language),
      translate("territory war 2",ws.api_language),
      translate("territory war 3",ws.api_language),
      translate("territory war 4",ws.api_language),
      translate("territory war 5",ws.api_language),
      translate("territory war 6",ws.api_language),
      translate("territory war 7",ws.api_language),
      translate("territory war 8",ws.api_language)
    );
  }
  writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).clearContent();
  writeSheet.getRange(1,1,1, headers.length).setValues([headers]);
  writeSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');

  //-->Build data to write to sheet
  var writeData = [];
  var tempGuild = [];
  tempGuild.push(
    guildData[0].profile.id,
    guildData[0].profile.name,
    guildData[0].member.filter(mem => mem.memberLevel === 4)[0].playerName,
    guildData[0].member.filter(mem => mem.memberLevel === 4)[0].playerId,
    guildData[0].profile.guildGalacticPower,
    guildData[0].profile.memberCount,
    '= ( ('+ (guildData[0].nextChallengesRefresh * 1000) +'/1000)/86400 ) + ( DATEVALUE("1-1-1970") - DATEVALUE("12-30-1899") )'
  );
  if(options.events){
    //-->Raid Data
    let raidData = "";
    guildData[0].recentRaidResult.forEach(raid => {
      raidData += raid.raidId + ":" + raid.identifier.campaignMissionId +",";
    });
    raidData = raidData.substring(0, raidData.length - 1);
    tempGuild.push(raidData);
    //-->Territory Battle data
    let tbData = "";
    guildData[0].profile.guildEventTracker.forEach(tb => {
      tbData += tb.definitionId + ":" + tb.completedStars + ",";
    });
    tbData = tbData.substring(0, tbData.length -1);
    tempGuild.push(tbData);
    //-->Territory War data
    let status = "";
    guildData[0].recentTerritoryWarResult.forEach(tw => {
      status = (Number(tw.score) > Number(tw.opponentScore)) ? "win" : "loss";
      tempGuild.push(translate("result:",ws.api_language) + translate(status,ws.api_language) + translate(",score:",ws.api_language) + tw.score + translate(",opponent:",ws.api_language) + tw.opponentScore + translate("power:",ws.api_language) + tw.power);
    });
  }
  writeData.push(tempGuild);

  //-->Write data to sheet
  if(currentData.filter(g => { return g[0] === guildData[0].profile.id })[0] !== undefined){
    writeSheet.getRange(currentMap[guildData[0].profile.id], 1,1,writeData[0].length).setValues(writeData);
  } else{
    writeSheet.getRange((writeSheet.getLastRow()+1),1,writeData.length, writeData[0].length).setValues(writeData);
  }

  //-->Format Sheet
  deleteUnusedCells(writeSheet);
  //writeSheet.sort(2,true);
  //-->Load member data
  updatePlayerDataSheet(settingSheetName,guildData);

  return guildData
}


/**
 * Creates or adds player data to the rosterHeroData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} playerData - The currently retrieved player data
 * @returns {object} playerData - Returns the player data to be used with other functions.
 */
function updateHeroRosterDataSheet(settingSheetName, playerData = null){
  const ws = new SheetMap(settingSheetName);
  const options = { 
    "stats": ws.option_hero_stats,
    "abilities": ws.option_hero_abilities,
    "tiers": ws.option_hero_tiers
  };
  //-->Create Sheet and designate write sheet
  var writeSheet
  if(ws.roster_hero_sheet === null){
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_rosterHeroSheet);
    SpreadsheetApp.flush();
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_rosterHeroSheet);
    writeSheet.setFrozenRows(1);
    writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');
  } else {
    writeSheet = ws.roster_hero_sheet;
  }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);

  //-->Get game data
  var isGuild = false;
  var ids = [];
  if(playerData === null){
    if(ws.input_playerID.toString().indexOf(",") > -1){
      ids = ws.input_playerID.toString().split(",");
    } else{
      ids = [ws.input_playerID.toString()];
    }
    playerData = api.fetchPlayers(ids,false,true);
  } else{
    if(playerData[0].hasOwnProperty("inviteStatus")){
      isGuild = true;
    }
  }
  //-->Build headers
  let headers = [
      translate("ally code",ws.api_language),
      translate("base id",ws.api_language),
      translate("rarity",ws.api_language),
      translate("level",ws.api_language),
      translate("gear level",ws.api_language),
      translate("relic level",ws.api_language),
      translate("galactic power",ws.api_language)
  ];
  if(options.stats){
    headers.push(
      translate("speed",ws.api_language),
      translate("health",ws.api_language),
      translate("protection",ws.api_language),
      translate("potency",ws.api_language),
      translate("tenacity",ws.api_language),
      translate("physical damage",ws.api_language),
      translate("physical critical chance",ws.api_language),
      translate("special damage",ws.api_language),
      translate("special critical chance",ws.api_language),
      translate("critical damage",ws.api_language),
      translate("accuracy",ws.api_language),
      translate("armor",ws.api_language),
      translate("resistance",ws.api_language),
      translate("critical avoidance",ws.api_language),
      translate("health steal",ws.api_language),
      translate("dodge",ws.api_language),
      translate("deflection",ws.api_language),
      translate("armor penetration",ws.api_language),
      translate("resistance penetration",ws.api_language)
    );
  }
  if(options.abilities){
    headers.push(
      translate("zeta ids",ws.api_language),
      translate("omicron ids",ws.api_language),
      translate("ultimate",ws.api_language)
    );
  }
  if(options.tiers){
    for(let i=1; i < 8; i++){
      headers.push(
        translate("ability_"+ i +"_id",ws.api_language),
        translate("ability_"+ i +"_tier",ws.api_language)
      );
    }
  }
  writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).clearContent();
  writeSheet.getRange(1,1,1, headers.length).setValues([headers]);

  //-->Build Player Data
  var writeData = [];
  var tempUnit = [];
  var removeList = [];
  var statOptions = {gameStyle: true, calcGP: true};
  var statCalc = new StatsAPI.StatCalculator();
  statCalc.setGameData();
  if(isGuild){
    ws.player_sheet.getRange(2,1,ws.player_sheet.getLastRow() + 1, ws.player_sheet.getLastColumn()).getValues().forEach(player => {
      if(playerData[0].profile.id === player[0]){
        removeList[player[2]] = player[2];
      }
    });
    statCalc.calcPlayerStats(playerData[0].member, statOptions);
    playerData[0].member.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.rosterUnit.forEach(unit => {
        if(unit.combatType === 1){
          tempUnit = [];
          tempUnit.push(
            player.allyCode,
            unit.baseId,
            unit.currentRarity,
            unit.currentLevel,
            unit.currentTier,
            (unit.relic.currentTier > 2) ? unit.relic.currentTier -2 : 0,
            unit.gp
          );
          if(options.stats){
            tempUnit.push(
              unit.stats.final["5"] || 0,
              unit.stats.final["1"] || 0,
              unit.stats.final["28"] || 0,
              unit.stats.final["17"] || 0,
              unit.stats.final["18"] || 0,
              unit.stats.final["6"] || 0,	
              unit.stats.final["14"] || 0,
              unit.stats.final["7"] || 0,
              unit.stats.final["15"] || 0,
              unit.stats.final["16"] || 0,
              unit.stats.final["37"] || 0,
              unit.stats.final["8"] || 0,
              unit.stats.final["9"] || 0,
              unit.stats.final["39"] || 0,	
              unit.stats.final["27"] || 0,
              unit.stats.final["12"] || 0, //Dodge
              unit.stats.final["13"] || 0, //Deflection
              unit.stats.final["10"] || 0, //Armor Pen
              unit.stats.final["11"] || 0,  //Resistance Pen
            );
          }
          if(options.abilities){
            tempUnit.push(
              unit.skills.filter(skill => { return skill.hasZeta }).map(skill => { return skill.id }).toString(),
              unit.skills.filter(skill => { return skill.hasOmicron }).map(skill => {return skill.id}).toString(),
              unit.purchasedAbilityId.toString()
            );
          }
          if(options.tiers){
            for(var a=0; a < 7; a++){
              if(unit.skills[a]){
                tempUnit.push(unit.skills[a].id,unit.skills[a].tier);
              }else{
                tempUnit.push("","");
              }
            }
          }
          writeData.push(tempUnit);

        }
      });
    });
  }else{
    statCalc.calcPlayerStats(playerData, statOptions);
    playerData.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.rosterUnit.forEach(unit => {
        if(unit.combatType === 1){
          tempUnit = [];
          tempUnit.push(
            player.allyCode,
            unit.baseId,
            unit.currentRarity,
            unit.currentLevel,
            unit.currentTier,
            (unit.relic.currentTier > 2) ? unit.relic.currentTier -2 : 0,
            unit.gp
          );
          if(options.stats){
            tempUnit.push(
              unit.stats.final["5"] || 0,
              unit.stats.final["1"] || 0,
              unit.stats.final["28"] || 0,
              unit.stats.final["17"] || 0,
              unit.stats.final["18"] || 0,
              unit.stats.final["6"] || 0,	
              unit.stats.final["14"] || 0,
              unit.stats.final["7"] || 0,
              unit.stats.final["15"] || 0,
              unit.stats.final["16"] || 0,
              unit.stats.final["37"] || 0,
              unit.stats.final["8"] || 0,
              unit.stats.final["9"] || 0,
              unit.stats.final["39"] || 0,	
              unit.stats.final["27"] || 0,
              unit.stats.final["12"] || 0, //Dodge
              unit.stats.final["13"] || 0, //Deflection
              unit.stats.final["10"] || 0, //Armor Pen
              unit.stats.final["11"] || 0,  //Resistance Pen
            );
          }
          if(options.abilities){
            tempUnit.push(
              unit.skills.filter(skill => { return skill.hasZeta }).map(skill => { return skill.id }).toString(),
              unit.skills.filter(skill => { return skill.hasOmicron }).map(skill => {return skill.id}).toString(),
              unit.purchasedAbilityId.toString()
            );
          }
          if(options.tiers){
            for(var a=0; a < 7; a++){
              if(unit.skills[a]){
                tempUnit.push(unit.skills[a].id,unit.skills[a].tier);
              }else{
                tempUnit.push("","");
              }
            }
          }
          writeData.push(tempUnit.slice());

        }
      });      
    });
  }

  //-->Delete old data
  deleteEntries(writeSheet,removeList);
  //-->Write player data to sheet
  if(writeData.length > 0){
    SpreadsheetApp.flush();
    appendData(writeSheet,writeData);
  }
  //-->Format Sheet
  deleteUnusedCells(writeSheet);
  //sortSheets(writeSheet,[0,1]);

  return playerData;
}


/**
 * Creates or adds player data to the rosterShipData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} playerData - The currently retrieved player data
 * @returns {object} playerData - Returns the player data to be used with other functions.
 */
function updateShipRosterDataSheet(settingSheetName, playerData = null){
  const ws = new SheetMap(settingSheetName);
  const options = { 
    "stats": ws.option_ship_stats,
    "tiers": ws.option_ship_tiers
  };
  //-->Create Sheet and designate write sheet
  var writeSheet
  if(ws.roster_ship_sheet === null){
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_rosterShipSheet);
    SpreadsheetApp.flush();
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_rosterShipSheet);
    writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');
    writeSheet.setFrozenRows(1);
  } else {
    writeSheet = ws.roster_ship_sheet;
  }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);

  //-->Get game data
  var isGuild = false;
  var ids = [];
  if(playerData === null){
    if(ws.input_playerID.toString().indexOf(",") > -1){
      ids = ws.input_playerID.toString().split(",");
    } else{
      ids = [ws.input_playerID.toString()];
    }
    playerData = api.fetchPlayers(ids,false,true);
  } else{
    if(playerData[0].hasOwnProperty("inviteStatus")){
      isGuild = true;
    }
  }
  //-->Build headers
  let headers = [
      translate("ally code",ws.api_language),
      translate("base id",ws.api_language),
      translate("rarity",ws.api_language),
      translate("level",ws.api_language),
      translate("galactic power",ws.api_language)
  ];
  if(options.stats){
    headers.push(
      translate("speed",ws.api_language),
      translate("health",ws.api_language),
      translate("protection",ws.api_language),
      translate("potency",ws.api_language),
      translate("tenacity",ws.api_language),
      translate("physical damage",ws.api_language),
      translate("physical critical chance",ws.api_language),
      translate("special damage",ws.api_language),
      translate("special critical chance",ws.api_language),
      translate("critical damage",ws.api_language),
      translate("accuracy",ws.api_language),
      translate("armor",ws.api_language),
      translate("resistance",ws.api_language),
      translate("critical avoidance",ws.api_language),
      translate("health steal",ws.api_language),
      translate("dodge",ws.api_language),
      translate("deflection",ws.api_language),
      translate("armor penetration",ws.api_language),
      translate("resistance penetration",ws.api_language)
    );
  }
  if(options.tiers){
    for(let i=1; i < 7; i++){
      headers.push(
        translate("ability_"+ i +"_id",ws.api_language),
        translate("ability_"+ i +"_tier",ws.api_language)
      );
    }
  }
  writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).clearContent();
  writeSheet.getRange(1,1,1, headers.length).setValues([headers]);

  //-->Build Player Data
  var writeData = [];
  var tempUnit = [];
  var removeList = [];
  var statOptions = {gameStyle: true, calcGP: true};
  var statCalc = new StatsAPI.StatCalculator();
  statCalc.setGameData();
  if(isGuild){
    ws.player_sheet.getRange(2,1,ws.player_sheet.getLastRow() + 1, ws.player_sheet.getLastColumn()).getValues().forEach(player => {
      if(playerData[0].profile.id === player[0]){
        removeList[player[2]] = player[2];
      }
    });
    statCalc.calcPlayerStats(playerData[0].member, statOptions);
    playerData[0].member.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.rosterUnit.forEach(unit => {
        if(unit.combatType === 2){
          tempUnit = [];
          tempUnit.push(
            player.allyCode,
            unit.baseId,
            unit.currentRarity,
            unit.currentLevel,
            unit.gp
          );
          if(options.stats){
            tempUnit.push(
              unit.stats.final["5"] || 0,
              unit.stats.final["1"] || 0,
              unit.stats.final["28"] || 0,
              unit.stats.final["17"] || 0,
              unit.stats.final["18"] || 0,
              unit.stats.final["6"] || 0,	
              unit.stats.final["14"] || 0,
              unit.stats.final["7"] || 0,
              unit.stats.final["15"] || 0,
              unit.stats.final["16"] || 0,
              unit.stats.final["37"] || 0,
              unit.stats.final["8"] || 0,
              unit.stats.final["9"] || 0,
              unit.stats.final["39"] || 0,	
              unit.stats.final["27"] || 0,
              unit.stats.final["12"] || 0, //Dodge
              unit.stats.final["13"] || 0, //Deflection
              unit.stats.final["10"] || 0, //Armor Pen
              unit.stats.final["11"] || 0,  //Resistance Pen
            );
          }
          if(options.tiers){
            for(var a=0; a < 6; a++){
              if(unit.skills[a]){
                tempUnit.push(unit.skills[a].id,unit.skills[a].tier);
              }else{
                tempUnit.push("","");
              }
            }
          }
          writeData.push(tempUnit);

        }
      });
    });
  }else{
    statCalc.calcPlayerStats(playerData, statOptions);
    playerData.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.rosterUnit.forEach(unit => {
        if(unit.combatType === 2){
          tempUnit = [];
          tempUnit.push(
            player.allyCode,
            unit.baseId,
            unit.currentRarity,
            unit.currentLevel,
            unit.gp
          );
          if(options.stats){
            tempUnit.push(
              unit.stats.final["5"] || 0,
              unit.stats.final["1"] || 0,
              unit.stats.final["28"] || 0,
              unit.stats.final["17"] || 0,
              unit.stats.final["18"] || 0,
              unit.stats.final["6"] || 0,	
              unit.stats.final["14"] || 0,
              unit.stats.final["7"] || 0,
              unit.stats.final["15"] || 0,
              unit.stats.final["16"] || 0,
              unit.stats.final["37"] || 0,
              unit.stats.final["8"] || 0,
              unit.stats.final["9"] || 0,
              unit.stats.final["39"] || 0,	
              unit.stats.final["27"] || 0,
              unit.stats.final["12"] || 0, //Dodge
              unit.stats.final["13"] || 0, //Deflection
              unit.stats.final["10"] || 0, //Armor Pen
              unit.stats.final["11"] || 0,  //Resistance Pen
            );
          }
          if(options.tiers){
            for(var a=0; a < 6; a++){
              if(unit.skills[a]){
                tempUnit.push(unit.skills[a].id,unit.skills[a].tier);
              }else{
                tempUnit.push("","");
              }
            }
          }
          writeData.push(tempUnit);

        }
      });      
    });
  }

  //-->Delete old data
  deleteEntries(writeSheet,removeList);
  //-->Write player data to sheet
  if(writeData.length > 0){
    appendData(writeSheet,writeData);
  }
  //-->Format Sheet
  deleteUnusedCells(writeSheet);
  //sortSheets(writeSheet,[0,1]);

  return playerData;
}


/**
 * Creates or adds mod data to the modData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} playerData - The currently retrieved player data
 * @returns {object} playerData - Returns the player data to be used with other functions.
 */
function updateModDataSheet(settingSheetName, playerData = null){
  const ws = new SheetMap(settingSheetName);
  const options = { 
    "rolls": ws.option_mod_rolls,
    "stats": ws.option_mod_stats
  };
  //-->Create Sheet and designate write sheet
  var writeSheet
  if(ws.mod_sheet === null){
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_modSheet);
    SpreadsheetApp.flush();
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_modSheet);
    writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');
    writeSheet.setFrozenRows(1);
  } else {
    writeSheet = ws.mod_sheet;
  }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);

  //-->Get game data
  var isGuild = false;
  var ids = [];
  if(playerData === null){
    if(ws.input_playerID.toString().indexOf(",") > -1){
      ids = ws.input_playerID.toString().split(",");
    } else{
      ids = [ws.input_playerID.toString()];
    }
    playerData = api.fetchPlayers(ids,false,true);
  } else{
    if(playerData[0].hasOwnProperty("inviteStatus")){
      isGuild = true;
    }
  }
  //-->Build headers
  let headers = [
      translate("ally code",ws.api_language),
      translate("base id",ws.api_language),
      translate("mod id",ws.api_language),
      translate("rarity",ws.api_language),
      translate("tier",ws.api_language),
      translate("level",ws.api_language),
      translate("slot",ws.api_language),
      translate("set",ws.api_language),
      translate("primary stat",ws.api_language),
      translate("primary amount",ws.api_language),
      translate("secondary stat A",ws.api_language),
      translate("A amount",ws.api_language),
      translate("A rolls",ws.api_language),
      translate("secondary stat B",ws.api_language),
      translate("B amount",ws.api_language),
      translate("B rolls",ws.api_language),
      translate("secondary stat C",ws.api_language),
      translate("C amount",ws.api_language),
      translate("C rolls",ws.api_language),
      translate("secondary stat D",ws.api_language),
      translate("D amount",ws.api_language),
      translate("D rolls",ws.api_language)
  ];
  if(options.rolls){
    headers.push(
      translate("A roll values",ws.api_language),
      translate("B roll values",ws.api_language),
      translate("C roll values",ws.api_language),
      translate("D roll values",ws.api_language)
    );
  }
  if(options.stats){
    headers.push(
      translate("critical chance",ws.api_language),
      translate("defense",ws.api_language),
      translate("defense %",ws.api_language),
      translate("health",ws.api_language),
      translate("health %",ws.api_language),
      translate("offense",ws.api_language),
      translate("offense %",ws.api_language),
      translate("protection",ws.api_language),
      translate("protection %",ws.api_language),
      translate("potency",ws.api_language),
      translate("speed",ws.api_language),
      translate("tenacity",ws.api_language),
      translate("search id",ws.api_language)
    );
  }
  writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).clearContent();
  writeSheet.getRange(1,1,1, headers.length).setValues([headers]);

  //-->Build Player Data
  var writeData = [];
  var tempMod = [];
  var removeList = [];
  var slotShape = {
    1:translate("Square",ws.api_language),2:translate("Arrow",ws.api_language),
    3:translate("Diamond",ws.api_language),4:translate("Triangle",ws.api_language),
    5:translate("Circle",ws.api_language),6:translate("Plus",ws.api_language)
  };
  var statName = {
    1:translate("Health",ws.api_language),
    5:translate("Speed",ws.api_language),
    16:translate("Critical Damage",ws.api_language),
    17:translate("Potency",ws.api_language),
    18:translate("Tenacity",ws.api_language),
    28:translate("Protection",ws.api_language),
    41:translate("Offense",ws.api_language),
    42:translate("Defense",ws.api_language),
    48:translate("Offense %",ws.api_language),
    49:translate("Defense %",ws.api_language),
    52:translate("Accuracy",ws.api_language),
    53:translate("Critical Chance",ws.api_language),
    54:translate("Critical Avoidance",ws.api_language),
    55:translate("Health %",ws.api_language),
    56:translate("Protection %",ws.api_language)
  };
  var setName = {
    1:statName[1],2:statName[41],3:statName[42],4:statName[5],
    5:statName[53],6:statName[16],7:statName[17],8:statName[18]
  };
  var statList = [ 0,0,0,0, 0,0,0,0, 0,0,0,0];
  var statPos = {
    53: 0, 42: 1, 49: 2, 1: 3, 55: 4,
    41: 5, 48: 6, 28: 7, 56: 8, 17: 9, 5: 10, 18: 11
  };
  var rollList = [];
  var searchData = [];

  if(isGuild){
    ws.player_sheet.getRange(2,1,ws.player_sheet.getLastRow() + 1, ws.player_sheet.getLastColumn()).getValues().forEach(player => {
      if(playerData[0].profile.id === player[0]){
        removeList[player[2]] = player[2];
      }
    });
    playerData[0].member.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.rosterUnit.forEach(unit => {
        if(unit.combatType === 1){
          unit.mods.forEach(mod => {
            tempMod = [];
            rollList = [];
            searchData = [];
            statList = [ 0,0,0,0, 0,0,0,0, 0,0,0,0];
            tempMod.push(
              player.allyCode,
              unit.baseId,
              mod.id,
              mod.pips,
              mod.tier,
              mod.level,
              slotShape[mod.slot],
              setName[mod.set],
              statName[mod.primaryStat.unitStat],
              mod.primaryStat.value
            );
            searchData.push(mod.slot,mod.set,mod.primaryStat.unitStat);
            for(let ss=0; ss < 4; ss++){
              if(mod.secondaryStat[ss] === undefined){
                tempMod.push("","","");
                rollList.push("");
                searchData.push("");
              } else {
                tempMod.push(statName[mod.secondaryStat[ss].unitStat], mod.secondaryStat[ss].value, mod.secondaryStat[ss].roll);
                statList[ statPos[mod.secondaryStat[ss].unitStat] ] = mod.secondaryStat[ss].value;
                rollList.push(mod.secondaryStat[ss].rollValues.toString());
                searchData.push(mod.secondaryStat[ss].unitStat);
              }
            }
            //Add roll values and stats
            if(options.rolls){
              rollList.forEach(roll => {
                tempMod.push(roll);
              })
            }
            if(options.stats){
              statList.forEach(stat => {
                tempMod.push(stat);
              });
              tempMod.push(getModSearchID(searchData));
            }
            writeData.push(tempMod);
          });
        }
      });
    });
  }else{
    playerData.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.rosterUnit.forEach(unit => {
        if(unit.combatType === 1){
          unit.mods.forEach(mod => {
            tempMod = [];
            rollList = [];
            searchData = [];
            statList = [ 0,0,0,0, 0,0,0,0, 0,0,0,0];
            tempMod.push(
              player.allyCode,
              unit.baseId,
              mod.id,
              mod.pips,
              mod.tier,
              mod.level,
              slotShape[mod.slot],
              setName[mod.set],
              statName[mod.primaryStat.unitStat],
              mod.primaryStat.value
            );
            searchData.push(mod.slot,mod.set,mod.primaryStat.unitStat);
            for(let ss=0; ss < 4; ss++){
              if(mod.secondaryStat[ss] === undefined){
                tempMod.push("","","");
                rollList.push("");
                searchData.push("");
              } else {
                tempMod.push(statName[mod.secondaryStat[ss].unitStat], mod.secondaryStat[ss].value, mod.secondaryStat[ss].roll);
                statList[ statPos[mod.secondaryStat[ss].unitStat] ] = mod.secondaryStat[ss].value;
                rollList.push(mod.secondaryStat[ss].rollValues.toString());
                searchData.push(mod.secondaryStat[ss].unitStat);
              }
            }
            //Add roll values and stats
            if(options.rolls){
              rollList.forEach(roll => {
                tempMod.push(roll);
              })
            }
            if(options.stats){
              statList.forEach(stat => {
                tempMod.push(stat);
              });
              tempMod.push(getModSearchID(searchData));
            }
            writeData.push(tempMod);
          });
        }
      });      
    });
  }

  //-->Delete old data
  deleteEntries(writeSheet,removeList);
  //-->Write player data to sheet
  if(writeData.length > 0){
    SpreadsheetApp.flush();
    appendData(writeSheet,writeData);
  }
  //-->Format Sheet
  //sortSheets(writeSheet,[0,1]);
  deleteUnusedCells(writeSheet);

  return playerData;
}


/**
 * Creates or adds mod data to the modData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} playerData - The currently retrieved player data
 * @returns {object} playerData - Returns the player data to be used with other functions.
 */
function updateDatacronDataSheet(settingSheetName, playerData = null){
  const ws = new SheetMap(settingSheetName);
  //-->Create Sheet and designate write sheet
  var writeSheet
  if(ws.datacron_sheet === null){
    SpreadsheetApp.flush();
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_datacronSheet);
    SpreadsheetApp.flush();
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_datacronSheet);
    writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');
    writeSheet.setFrozenRows(1);
  } else {
    writeSheet = ws.datacron_sheet;
  }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);

  //-->Get game data
  var isGuild = false;
  var ids = [];
  if(playerData === null){
    if(ws.input_playerID.toString().indexOf(",") > -1){
      ids = ws.input_playerID.toString().split(",");
    } else{
      ids = [ws.input_playerID.toString()];
    }
    playerData = api.fetchPlayers(ids,false,true);
  } else{
    if(playerData[0].hasOwnProperty("inviteStatus")){
      isGuild = true;
    }
  }
  //-->Build headers
  let headers = [
      translate("ally code",ws.api_language),
      translate("datacron id",ws.api_language),
      translate("datacron set",ws.api_language),
      translate("max level",ws.api_language)
  ];
  for(let i=1; i < 10; i++){
    headers.push(translate("level "+i+ " target",ws.api_language), translate("level "+i+ " value",ws.api_language));
  }
  writeSheet.getRange(1,1,1,writeSheet.getMaxColumns()).clearContent();
  writeSheet.getRange(1,1,1, headers.length).setValues([headers]);

  //-->Build Player Data
  var writeData = [];
  var tempCron = [];
  var removeList = [];

  if(isGuild){
    ws.player_sheet.getRange(2,1,ws.player_sheet.getLastRow() + 1, ws.player_sheet.getLastColumn()).getValues().forEach(player => {
      if(playerData[0].profile.id === player[0]){
        removeList[player[2]] = player[2];
      }
    });
    playerData[0].member.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.datacron.forEach(datacron => {
        tempCron = [];
        tempCron.push(
          player.allyCode,
          datacron.id,
          datacron.setName,
          datacron.maxTier
        );
        for(let lv=0; lv < 9; lv++){
          if(datacron.affix[lv] === undefined){
            tempCron.push("","");
          } else {
            tempCron.push(
              datacron.affix[lv].targetNameKey, 
              (datacron.affix[lv].statValue === 0) ? datacron.affix[lv].abilityDescKey : datacron.affix[lv].statValue
            );
          }
        }
        writeData.push(tempCron);
      });
    });
  }else{
    playerData.forEach(player => {
      removeList[player.allyCode.toString()] = player.allyCode;
      player.datacron.forEach(datacron => {
        tempCron = [];
        tempCron.push(
          player.allyCode,
          datacron.id,
          datacron.setName,
          datacron.maxTier
        );
        for(let lv=0; lv < 9; lv++){
          if(datacron.affix[lv] === undefined){
            tempCron.push("","");
          } else {
            tempCron.push(
              datacron.affix[lv].targetNameKey, 
              (datacron.affix[lv].abilityDescKey !== "") ? datacron.affix[lv].abilityDescKey : datacron.affix[lv].statValue
            );
          }
        }
        writeData.push(tempCron);
      });
    });
  }

  //-->Delete old data
  deleteEntries(writeSheet,removeList);
  //-->Write player data to sheet
  if(writeData.length > 0){
    appendData(writeSheet,writeData);
  }
  //-->Format Sheet
  //sortSheets(writeSheet,[0,2,3]);
  deleteUnusedCells(writeSheet);

  return playerData;
}
