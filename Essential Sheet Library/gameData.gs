/*
Functions found here
- loadGameData(settingSheetName)
- updateUnitDataSheet(settingSheetName, gameData, append)
- updateBaseDataSheet(settingSheetName, gameData, append)
- updateHeroAbilitySheet(settingSheetName, gameData, append)
- updateShipAbilitySheet(settingSheetName, gameData, append)
- updateGearSheet(settingSheetName, gameData, append)
- updateRelicSheet(settingSheetName, gameData, append)
*/

/**
 * Builds the non-player data sheets with the selected options from the essentialSettings tab.
 * @param {string} settingSheetName - The name of the settings sheet
 */
function loadGameData(settingSheetName){
  //Verify sheets to create
  const ws = new SheetMap(settingSheetName);
  const create = {
    "unitData": { "include": ws.create_units, "sheetName": ws.name_unitSheet },
    "heroBaseData": { "include": ws.create_heroBase, "sheetName": ws.name_baseSheet },
    "heroAbilityData": { "include": ws.create_heroAbility, "sheetName": ws.name_heroAbilitySheet },
    "shipAbilityData": { "include": ws.create_shipAbility, "sheetName": ws.name_shipAbilitySheet },
    "gearData": { "include": ws.create_gear, "sheetName": ws.name_gearSheet },
    "relicData": { "include": ws.create_relic, "sheetName": ws.name_relicSheet }
  }
  var gameData = null;
  //Check if sheets exist
  for(let key in create){
    if(create[key].include){
      //Build Sheet
      switch(key){
        case "unitData":
          gameData = updateUnitDataSheet(settingSheetName, gameData);
          break;
        case "heroBaseData":
          gameData = updateBaseDataSheet(settingSheetName, gameData);
          break;
        case "heroAbilityData":
          gameData = updateHeroAbilitySheet(settingSheetName, gameData);
          break;
        case "shipAbilityData":
          gameData = updateShipAbilitySheet(settingSheetName, gameData);
          break;
        case "gearData":
          gameData = updateGearSheet(settingSheetName, gameData);
          break;
        case "relicData":
          gameData = updateRelicSheet(settingSheetName, gameData);
          break;
      }
    }
  }
}

/**
 * Updates the unitData sheet using the chosen options from the essentialSettings sheet.
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} gameData - The currently retrieved gameData collections
 * @param {boolean} append - Flag to add new data changes instead of replacing everything
 * @returns {object} gameData - Returns the gameData to be used with other functions.
 */
function updateUnitDataSheet(settingSheetName, gameData = null, append = false){
  const ws = new SheetMap(settingSheetName);
  //-->Create Sheet and designate write sheet
  var writeSheet
  var currentData = [];
  if(ws.unit_sheet === null){
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_unitSheet);
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_unitSheet);
  } else {
    writeSheet = ws.unit_sheet;
    currentData = (writeSheet.getRange("A1").getValue() === "") ? [] : writeSheet.getRange(1,1,writeSheet.getLastRow(),writeSheet.getLastColumn()).getValues();
  }
  //-->Set dynamic array indexes
  const options = { "images": ws.option_unit_images };
  var nameIndx = (options.images) ? 2 : 1;
  var baseIndx = (options.images) ? 1 : 0;
  var typeIndx = (options.images) ? 3 : 2;
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  if(gameData !== null){
    if(!gameData.hasOwnProperty("units")){ gameData['units'] = api.fetchData(['units']).units; }
    if(!gameData.hasOwnProperty("category")){ gameData['category'] = api.fetchData(['category']).category; }
    if(!gameData.hasOwnProperty("localization")){ gameData['localization'] = api.fetchLocalization(); }
  }else {
    gameData = api.fetchData(['units', 'category']);
    gameData['localization'] = api.fetchLocalization();
  }
  var categoryMap = [];
  for(let i=0; i < gameData["category"].length; i++){
    categoryMap[ gameData.category[i].id ] = i;
  }
  //-->Get any current data
  var currentMap = [];
  if(currentData.length > 0){
    for(var i=1;i < currentData.length;i++){
      currentMap[ currentData[i][(baseIndx)]] = i;
    }
  }
  var unitData = [];
  var writeData = [];
  var units = [];

  //Build Headers
  let headers = [
      translate("base id",ws.api_language),
      translate("name",ws.api_language),
      translate("combat type",ws.api_language),
      translate("categories",ws.api_language),
      translate("crew|ship",ws.api_language)
  ];
  if(options.images){
    headers.unshift(translate("image",ws.api_language));
  }
  writeData.push(headers);
  if(append && currentData.length > 0){
    for(let col=writeData[0].length; col < currentData[0].length; col++){
      writeData[0].push(currentData[0][col]);
    }
  }

  //Build data to write
  let combat, name, id, image 
  let categories = ""; 
  let crew = [];
  let getShip = [];
  var indx
  gameData["units"].forEach(unit => {
    if(unit.maxRarity === unit.rarity && unit.obtainable === true && unit.obtainableTime === "0"){
      categories = [];
      crew = [];
      combat = unit.combatType;
      name = gameData["localization"][unit.nameKey];
      id = unit.baseId;
      image = '=IMAGE("https://game-assets.swgoh.gg/' + unit.thumbnailName +'.png",1)';
      unit.crew.forEach(crw => {
        getShip[crw.unitId] = unit.baseId;
        crew.push(crw.unitId);
      });
      crew = crew.toString();
      unit.categoryId.forEach(cat => {
        indx = categoryMap[ cat ];
        if(gameData["category"][indx].visible){
          categories.push(gameData["localization"][ gameData["category"][indx].descKey ]);
        }
      });
      categories = categories.toString();
      if(options.images){
        units = [image, id, name,combat,categories,crew];
      }else{
        units = [id,name,combat,categories,crew];
      }
      if(append){
        indx = currentMap[id];
        for(let col=units.length; col < currentData[0].length; col++){
          if(indx === undefined){
            units.push("");
          }else{
            units.push(currentData[indx][col]);
          }
        }
      }
      unitData.push(units);
    }
  });
  //Go back and apply ships to characters
  unitData.forEach(unit => {
    if( getShip.hasOwnProperty( unit[baseIndx] ) ){
      unit[(baseIndx+4)] = getShip[unit[baseIndx]];
    }
  });

  //Sort Data
  unitData.sort((a,b) =>{
    if(a[typeIndx] === b[typeIndx]){
      return a[nameIndx] < b[nameIndx] ? -1 : 1;
    }else{
      return a[typeIndx] < b[typeIndx] ? -1 : 1;
    }
  });
  writeData = [].concat(writeData,unitData);
  //Write Sheet
  if(writeSheet.getRange("A1").getValue() !== ""){
    writeSheet.getRange(1,1, writeSheet.getLastRow(),writeSheet.getLastColumn()).clear();
  }
  writeSheet.getRange(1,1, writeData.length, writeData[0].length).setValues(writeData);
  //Format sheet
  writeSheet.getRange(1,1,1,writeData[0].length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc'); 
  deleteUnusedCells(writeSheet);
  writeSheet.setFrozenRows(1);
  //writeSheet.autoResizeColumns(1,writeData[0].length);

  return gameData; //Return data in case calling other functions that need it.
}


/**
 * Updates the heroBaseData sheet using the chosen options from the essentialSettings sheet.
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} gameData - The currently retrieved gameData collections
 * @param {boolean} append - Flag to add new data changes instead of replacing everything
 * @returns {object} gameData - Returns the gameData to be used with other functions.
 */
function updateBaseDataSheet(settingSheetName, gameData = null, append = false){
  const ws = new SheetMap(settingSheetName);
  //-->Create Sheet and designate write sheet
  var currentData = [];
  var writeSheet
    if(ws.base_sheet === null){
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_baseSheet);
      writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_baseSheet);
    } else {
      writeSheet = ws.base_sheet;
      currentData = (writeSheet.getRange("A1").getValue() === "") ? [] : writeSheet.getRange(1,1,writeSheet.getLastRow(),writeSheet.getLastColumn()).getValues();
    }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  if(gameData !== null){
    if(!gameData.hasOwnProperty("units")){ gameData['units'] = api.fetchData(['units']).units; }
    if(!gameData.hasOwnProperty("localization")){ gameData['localization'] = api.fetchLocalization(); }
  }else {
    gameData = api.fetchData(['units']);
    gameData['localization'] = api.fetchLocalization();
  }
  const options = { "stats": ws.option_base_stats };
  const statCalc = new StatsAPI.StatCalculator();
  statCalc.setGameData();
  const relicCount = Object.keys(statCalc.unitData.ADMIRALACKBAR.relic).length + 2;
  var statOptions = {gameStyle: false, calcGP: true, language: statCalc.lang, withoutModCalc: true, useValues:{}};
  var unitData = [];
  var writeData = [];


  //Build Headers
  let headers = [
      translate("base id",ws.api_language),
      translate("name",ws.api_language),
      translate("gear level",ws.api_language),
      translate("relic level",ws.api_language),
      translate("equipment",ws.api_language)
  ];
  if(options.stats){
    headers.push(
      translate("speed",ws.api_language),
      translate("health",ws.api_language),
      translate("protection",ws.api_language),
      translate("potency",ws.api_language),
      translate("tenacity",ws.api_language),
      translate("physical damage",ws.api_language),
      translate("physical critical rating",ws.api_language),
      translate("special damage",ws.api_language),
      translate("special critical rating",ws.api_language),
      translate("critical damage",ws.api_language),
      translate("accuracy",ws.api_language),
      translate("armor rating",ws.api_language),
      translate("restistance rating",ws.api_language),
      translate("critical avoidance",ws.api_language),
      translate("health steal",ws.api_language),
      translate("dodge rating",ws.api_language),
      translate("deflection rating",ws.api_language),
      translate("armor penetration",ws.api_language),
      translate("resistance penetration",ws.api_language)
    );
  }
  writeData.push(headers);
  if(append && currentData.length > 0){
    for(let col=writeData[0].length; col < currentData[0].length; col++){
      writeData[0].push(currentData[0][col]);
    }
  }


  //Build data to write
  var tempUnit
  gameData["units"].forEach(unit => {
    if(unit.rarity === 7 && unit.combatType === 1 && unit.obtainable === true && unit.obtainableTime === "0"){
      for(var gear=1; gear < 14;gear++){
        tempUnit = currentData.filter(function(u) { return u[0] === unit.baseId && u[2] === gear; });
        unitData = [];
        unitData.push(
          unit.baseId,
          gameData.localization[unit.nameKey],
          gear,
          0,
          (gear === 13) ? "RELIC_1" : unit.unitTier[gear].equipmentSet.toString()
        );
        if(options.stats){
          statOptions.useValues = {
            char: { // used when calculating character stats
              rarity: 7,
              level: 85,
              gear: gear,
              equipped: "none", // See Below
              relic: 1 // 1='locked', 2='unlocked', 3=R1, 4=R2, ...9=R7
            }
          };
          stats = statCalc.calcCharStats(unit.baseId,statOptions);
          unitData.push(
            stats.base["Speed"],
            stats.base["Health"],
            stats.base["Protection"] || 0,
            stats.base["Potency"] || 0,
            stats.base["Tenacity"],
            stats.base["Physical Damage"],
            stats.base["Physical Critical Chance"],
            stats.base["Special Damage"],
            stats.base["Special Critical Chance"],
            stats.base["Critical Damage"],
            (((stats.base["Physical Accuracy"])/1200)) || 0,
            stats.base["Armor"],
            stats.base["Resistance"],
            (((stats.base["Physical Critical Avoidance"])/2400)) || 0,
            stats.base["Health Steal"] || 0,
            stats.base["Dodge Chance"] || 0.02,
            stats.base["Deflection Chance"] || 0.02,
            stats.base["Armor Penetration"] || 0,
            stats.base["Resistance Penetration"] || 0
          );
        }
        //Add existing data
        if(append && currentData.length > 0){
          for(let col=unitData.length; col < currentData[0].length; col++){
            if(tempUnit[0] === undefined){
              unitData.push("");
            }else{
              unitData.push(tempUnit[0][col]);
            }
          }
        }
        writeData.push(unitData);
      }
      //=>Relic Stats
      for(var relic=3; relic < (relicCount+1);relic++){
        tempUnit = currentData.filter(function(u) { return u[0] === unit.baseId && u[3] === relic; });
        unitData = [];
        unitData.push(
          unit.baseId,
          gameData.localization[unit.nameKey],
          13,
          (relic-2),
          (relic === relicCount) ? "" : "RELIC_" + (relic-1)
        );
        if(options.stats){
          statOptions.useValues = {
            char: { // used when calculating character stats
              rarity: 7,
              level: 85,
              gear: 13,
              equipped: "none", // See Below
              relic: relic // 1='locked', 2='unlocked', 3=R1, 4=R2, ...9=R7
            }
            };
          stats = statCalc.calcCharStats(unit.baseId,statOptions);
          unitData.push(
            stats.base["Speed"],
            stats.base["Health"],
            stats.base["Protection"] || 0,
            stats.base["Potency"] || 0,
            stats.base["Tenacity"],
            stats.base["Physical Damage"],
            stats.base["Physical Critical Chance"],
            stats.base["Special Damage"],
            stats.base["Special Critical Chance"],
            stats.base["Critical Damage"],
            (((stats.base["Physical Accuracy"])/1200)) || 0,
            stats.base["Armor"],
            stats.base["Resistance"],
            (((stats.base["Physical Critical Avoidance"])/2400)) || 0,
            stats.base["Health Steal"] || 0,
            stats.base["Dodge Chance"] || 0.02,
            stats.base["Deflection Chance"] || 0.02,
            stats.base["Armor Penetration"] || 0,
            stats.base["Resistance Penetration"] || 0
          );
        }
        //Add existing data
        if(append && currentData.length > 0){
          for(let col=unitData.length; col < currentData[0].length; col++){
            if(tempUnit[0] === undefined){
              unitData.push("");
            }else{
              unitData.push(tempUnit[0][col]);
            }
          }
        }
        writeData.push(unitData);
      }
    }
  });

  //Clear Sheet
  if(writeSheet.getRange("A1").getValue() !== ""){
    writeSheet.clearContents().clear()
  }
  //Write Sheet
  writeSheet.getRange(1,1, writeData.length, writeData[0].length).setValues(writeData);
  //Format sheet
  writeSheet.getRange(1,1,1,writeData[0].length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');  
  deleteUnusedCells(writeSheet);
  //writeSheet.autoResizeColumns(1,writeData[0].length);
  writeSheet.getRange("E2:E").setNumberFormat('@STRING@');
  writeSheet.setFrozenRows(1);
  writeSheet.sort(1, true);

  return gameData; //Return data in case calling other functions that need it.
}


/**
 * Updates the heroAbilityData sheet using the chosen options from the essentialSettings sheet.
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} gameData - The currently retrieved gameData collections
 * @param {boolean} append - Flag to add new data changes instead of replacing everything
 * @returns {object} gameData - Returns the gameData to be used with other functions.
 */
function updateHeroAbilitySheet(settingSheetName, gameData = null, append = false){
  const ws = new SheetMap(settingSheetName);
  //-->Create Sheet and designate write sheet
  var writeSheet
  var currentData = [];
  if(ws.hero_ability_sheet === null){
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_heroAbilitySheet);
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_heroAbilitySheet);
  } else {
    writeSheet = ws.hero_ability_sheet;
    currentData = (writeSheet.getRange("A1").getValue() === "") ? [] : writeSheet.getRange(1,1,writeSheet.getLastRow(),writeSheet.getLastColumn()).getValues();
  }
  //-->Set dynamic array indexes
  const options = { "images": ws.option_ability_images, "tiers": ws.option_ability_tiers };
  const idCol = (options.images) ? 2 : 1;
  const unitCol = (options.images) ? 4 : 3;
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  if(gameData !== null){
    if(!gameData.hasOwnProperty("units")){ gameData['units'] = api.fetchData(['units']).units; }
    if(!gameData.hasOwnProperty("ability")){ gameData['ability'] = api.fetchData(['ability']).ability; }
    if(!gameData.hasOwnProperty("skill")){ gameData['skill'] = api.fetchData(['skill']).skill; }
    if(!gameData.hasOwnProperty("recipe")){ gameData['recipe'] = api.fetchData(['recipe']).recipe; }
    if(!gameData.hasOwnProperty("material")){ gameData['material'] = api.fetchData(['material']).material; }
    if(!gameData.hasOwnProperty("localization")){ gameData['localization'] = api.fetchLocalization(); }
  }else {
    gameData = api.fetchData(['units', 'ability', 'skill', 'recipe','material']);
    gameData['localization'] = api.fetchLocalization();
  }
  //==>Build Maps
  var skillMap = [];
  for(let sk=0; sk < gameData["skill"].length; sk++){
    skillMap[gameData["skill"][sk]["id"]] = sk;
  }
  var recipeMap = [];
  for(let r=0; r < gameData["recipe"].length; r++){
    recipeMap[gameData["recipe"][r]["id"]] = r;
  }
  var abilityMap = [];
  for(let r=0; r < gameData["ability"].length; r++){
    abilityMap[gameData["ability"][r]["id"]] = r;
  }  
  var materialMap = [];
  for(let r=0; r < gameData["material"].length; r++){
    materialMap[gameData["material"][r]["id"]] = r;
  }  

  //Build Headers
  var writeData = [];
  let headers = [
      translate("ability ID", ws.api_language),
      translate("ability name",ws.api_language),
      translate("unit base id",ws.api_language),
      translate("max tier",ws.api_language),
      translate("omega tier",ws.api_language),
      translate("zeta tier",ws.api_language),
      translate("omicron tier",ws.api_language),
      translate("omicron area",ws.api_language)
  ];
  if(options.images){
    headers.unshift(translate("image",ws.api_language));
  }
  if(options.tiers){
    for(let i=1; i < 10; i++){
      headers.push(
        translate("tier " + i + " description", ws.api_language),
        translate("tier " + i + " mats", ws.api_language),
        translate("tier " + i + " requirements", ws.api_language)
      );
    }
  }
  writeData.push(headers.slice());
  if(append && currentData.length > 0){
    for(let col=writeData[0].length; col < currentData[0].length; col++){
      writeData[0].push(currentData[0][col]);
    }
  }

  //=Write Data
  var abilityData = [];
  var tierData = [];
  var trGear = translate("Gear",ws.api_language);
  var trRelic = translate("Relic",ws.api_language);
  var trUnitLv = translate("Unit Level",ws.api_language);
  gameData["units"].forEach(unit => {
    if(unit.rarity === 7 && unit.combatType === 1 && unit.obtainable === true && unit.obtainableTime === "0"){
      let abID,abName,maxTier,omegaTier,omiTier,omiArea,zetaTier,image,tierMat,tierReq,sIndx,aIndx
      unit.skillReference.forEach(skill => {
        abilityData = [];
        tierData = [];
        omegaTier = 0;
        zetaTier = 0;
        omiTier = 0;
        abID = skill.skillId;
        sIndx = skillMap[skill.skillId];
        aIndx = abilityMap[gameData["skill"][sIndx].abilityReference];
        abName = gameData["localization"][gameData["ability"][aIndx].nameKey];
        image = '=IMAGE("https://game-assets.swgoh.gg/' + gameData["ability"][aIndx].icon + '.png",1)';
        tierReq = (skill.requiredRelicTier > 2) ? trRelic + ":" + (skill.requiredRelicTier - 2) : trGear + ":" + skill.requiredTier;
        maxTier = gameData["skill"][sIndx].tier.length + 1;
        omiArea = getOmicronArea(gameData["skill"][sIndx].omicronMode);
        //=> Set tier data
        for(let i=0; i < maxTier; i++ ){
          let tIndx = i - 1;
          if(i === 0){
            tierData.push([
              gameData["localization"][gameData["ability"][aIndx].descKey],
              "",
              tierReq
            ]);
          } else {
            let rIndx = recipeMap[gameData["skill"][sIndx].tier[tIndx].recipeId];
            tierMat = "";
            gameData["recipe"][rIndx].ingredients.forEach(mat => {
              let mIndx = (materialMap.hasOwnProperty(mat.id)) ? materialMap[mat.id] : null;
              let matName = (mat.id === "GRIND") ? gameData["localization"]["Shared_Currency_Grind"] : gameData["localization"][gameData["material"][mIndx].nameKey];
              tierMat +=  matName + ":" + mat.maxQuantity + ",";
              switch(mat.id){
                case "ability_mat_D":
                  if(omegaTier === 0){ omegaTier = i + 1; }
                  break;
                case "ability_mat_E":
                  if(zetaTier === 0){ zetaTier = i + 1; }
                  break;
                case "ability_mat_F":
                  if(omiTier === 0){ omiTier = i + 1; }
                  break;
              }
            }); //Shared_Currency_ShipGrind
            tierMat = tierMat.substring(0, tierMat.length-1);
            tierData.push([
              gameData["localization"][gameData["ability"][aIndx].tier[tIndx].descKey],
              tierMat,
              trUnitLv + ":" + gameData["skill"][sIndx].tier[tIndx].requiredUnitLevel
            ]);
          }
        }
        abilityData.push(abID,abName,unit.baseId,maxTier,omegaTier,zetaTier,omiTier,omiArea);
        if(options.images){
          abilityData.unshift(image);
        }
        if(options.tiers){
          tierData.forEach(tier => {
            abilityData.push(tier[0],tier[1],tier[2]);
          });
          for(let col=abilityData.length; col < headers.length; col++){
            abilityData.push("");
          }
        }
        if(append  && currentData.length > 0){
          if(abilityData.length !== writeData[0].length){
            let tempAbility = currentData.filter(ability => { return ability[idCol-1] === abID; })[0];
            for(let col=abilityData.length; col < currentData[0].length; col++){
              if(tempAbility === undefined){
                abilityData.push("");
              } else{
                abilityData.push(tempAbility[col]);
              }
            }
          }
        }
        writeData.push(abilityData.slice());
      });
      //=>Add Ultimates
      if(unit.legend){
        let abRecipe
        abilityData = [];
        unit.limitBreakRef.forEach(ability => {
          if(ability.powerAdditiveTag === "ultimate"){
            abID = ability.abilityId;
            abRecipe = ability.unlockRecipeId;
            tierReq = (ability.requiredRelicTier > 2) ? trRelic + ":" + (ability.requiredRelicTier - 2) : trGear + ":" + ability.requiredTier;
          }
        });
        unit.uniqueAbilityRef.forEach(ability => {
          if(ability.powerAdditiveTag === "ultimate"){
            abID = ability.abilityId;
            abRecipe = ability.unlockRecipeId;
            tierReq = (ability.requiredRelicTier > 2) ? trRelic + ":" + (ability.requiredRelicTier - 2) : trGear + ":" + ability.requiredTier;
          }
        });
        abName = gameData["localization"][gameData["ability"][abilityMap[abID]].nameKey];
        abDesc = gameData["localization"][gameData["ability"][abilityMap[abID]].descKey];
        image = '=IMAGE("https://game-assets.swgoh.gg/' + gameData["ability"][abilityMap[abID]].icon + '.png",1)';
        maxTier = gameData["ability"][abilityMap[abID]].tier.length + 1;
        abilityData.push(abID,abName,unit.baseId,maxTier,0,0,0,"");
        if(options.images){
          abilityData.unshift(image);
        }
        if(options.tiers){
          let rIndx = recipeMap[abRecipe];
          gameData["recipe"][rIndx].ingredients.forEach(mat => {
            tierMat = "";
            let mIndx = (materialMap.hasOwnProperty(mat.id)) ? materialMap[mat.id] : null;
            let matName = (mat.id === "GRIND") ? gameData["localization"]["Shared_Currency_Grind"] : gameData["localization"][gameData["material"][mIndx].nameKey];
            tierMat +=  matName + ":" + mat.maxQuantity + ",";
          });
          tierMat = tierMat.substring(0, tierMat.length-1);
          abilityData.push(abDesc,tierMat,tierReq);
          for(let col=abilityData.length; col < headers.length; col++){
            abilityData.push("");
          }
        }
        if(append  && currentData.length > 0){
          if(abilityData.length !== writeData[0].length){
            let tempAbility = currentData.filter(ability => { return ability[idCol-1] === abID; })[0];
            for(let col=abilityData.length; col < currentData[0].length; col++){
              if(tempAbility === undefined){
                abilityData.push("");
              } else{
                abilityData.push(tempAbility[col]);
              }
            }
          }
        }
        writeData.push(abilityData.slice());
      }
    }
  });

  //Clear Sheet
  if(writeSheet.getRange("A1").getValue() !== ""){
    writeSheet.clearContents().clear()
  }
  //Write Sheet
  writeSheet.getRange(1,1, writeData.length, writeData[0].length).setValues(writeData);
  //Format sheet
  writeSheet.getRange(1,1,1,writeData[0].length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');  
  deleteUnusedCells(writeSheet);
  //writeSheet.autoResizeColumns(1,writeData[0].length);
  writeSheet.setFrozenRows(1);
  writeSheet.sort(idCol, true);
  writeSheet.sort(unitCol, true);

  return gameData; //Return data in case calling other functions that need it.
}


/**
 * Updates the shipAbilityData sheet using the chosen options from the essentialSettings sheet.
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} gameData - The currently retrieved gameData collections
 * @param {boolean} append - Flag to add new data changes instead of replacing everything
 * @returns {object} gameData - Returns the gameData to be used with other functions.
 */
function updateShipAbilitySheet(settingSheetName, gameData = null, append = false){
  const ws = new SheetMap(settingSheetName);
  //-->Create Sheet and designate write sheet
  var writeSheet
  var currentData = [];
  if(ws.ship_ability_sheet === null){
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_shipAbilitySheet);
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_shipAbilitySheet);
  } else {
    writeSheet = ws.ship_ability_sheet;
    currentData = (writeSheet.getRange("A1").getValue() === "") ? [] : writeSheet.getRange(1,1,writeSheet.getLastRow(),writeSheet.getLastColumn()).getValues();
  }
  //-->Set dynamic array indexes
  const options = { "images": ws.option_ability_images, "tiers": ws.option_ability_tiers };
  const idCol = (options.images) ? 2 : 1;
  const unitCol = (options.images) ? 4 : 3;
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  if(gameData !== null){
    if(!gameData.hasOwnProperty("units")){ gameData['units'] = api.fetchData(['units']).units; }
    if(!gameData.hasOwnProperty("ability")){ gameData['ability'] = api.fetchData(['ability']).ability; }
    if(!gameData.hasOwnProperty("skill")){ gameData['skill'] = api.fetchData(['skill']).skill; }
    if(!gameData.hasOwnProperty("recipe")){ gameData['recipe'] = api.fetchData(['recipe']).recipe; }
    if(!gameData.hasOwnProperty("material")){ gameData['material'] = api.fetchData(['material']).material; }
    if(!gameData.hasOwnProperty("localization")){ gameData['localization'] = api.fetchLocalization(); }
  }else {
    gameData = api.fetchData(['units', 'ability', 'skill', 'recipe','material']);
    gameData['localization'] = api.fetchLocalization();
  }
  //==>Build Maps
  var skillMap = [];
  for(let sk=0; sk < gameData["skill"].length; sk++){
    skillMap[gameData["skill"][sk]["id"]] = sk;
  }
  var recipeMap = [];
  for(let r=0; r < gameData["recipe"].length; r++){
    recipeMap[gameData["recipe"][r]["id"]] = r;
  }
  var abilityMap = [];
  for(let r=0; r < gameData["ability"].length; r++){
    abilityMap[gameData["ability"][r]["id"]] = r;
  }  
  var materialMap = [];
  for(let r=0; r < gameData["material"].length; r++){
    materialMap[gameData["material"][r]["id"]] = r;
  }  

  //Build Headers
  var writeData = [];
  let headers = [
      translate("ability ID", ws.api_language),
      translate("ability name",ws.api_language),
      translate("unit base id",ws.api_language),
      translate("max tier",ws.api_language),
      translate("crew skill",ws.api_language)
  ];
  if(options.images){
    headers.unshift(translate("image",ws.api_language));
  }
  if(options.tiers){
    for(let i=2; i < 9; i++){
      headers.push(
        translate("tier " + i + " description", ws.api_language),
        translate("tier " + i + " mats", ws.api_language),
        translate("tier " + i + " requirements", ws.api_language)
      );
    }
  }
  writeData.push(headers.slice());
  if(append && currentData.length > 0){
    for(let col=writeData[0].length; col < currentData[0].length; col++){
      writeData[0].push(currentData[0][col]);
    }
  }

  //=Write Data
  var skillList = [];
  var tierData = [];
  var trUnitLv = translate("Unit Level",ws.api_language);
  var trCrewGearLv = translate("Crew Gear Level",ws.api_language);
  gameData["units"].forEach(unit => {
    if(unit.rarity === 7 && unit.combatType === 2 && unit.obtainable === true && unit.obtainableTime === "0"){
      let abID,abName,maxTier,image,tierMat,sIndx,aIndx
      unit.skillReference.forEach(skill => {
        abilityData = [];
        tierData = [];
        abID = skill.skillId;
        sIndx = skillMap[skill.skillId];
        aIndx = abilityMap[gameData["skill"][sIndx].abilityReference];
        abName = gameData["localization"][gameData["ability"][aIndx].nameKey];
        image = '=IMAGE("https://game-assets.swgoh.gg/' + gameData["ability"][aIndx].icon + '.png",1)';
        maxTier = gameData["skill"][sIndx].tier.length + 1;
        //=> Set tier data
        for(let i=0; i < (maxTier-1); i++ ){
          let tIndx = i;
          let rIndx = recipeMap[gameData["skill"][sIndx].tier[tIndx].recipeId];
          tierMat = "";
          gameData["recipe"][rIndx].ingredients.forEach(mat => {
            let mIndx = (materialMap.hasOwnProperty(mat.id)) ? materialMap[mat.id] : null;
            let matName = (mat.id === "SHIP_GRIND") ? gameData["localization"]["Shared_Currency_ShipGrind"] : gameData["localization"][gameData["material"][mIndx].nameKey];
            tierMat +=  matName + ":" + mat.maxQuantity + ",";
          });
          tierMat = tierMat.substring(0, tierMat.length-1);
          tierData.push([
            gameData["localization"][gameData["ability"][aIndx].tier[tIndx].descKey],
            tierMat,
            trUnitLv + ":" + gameData["skill"][sIndx].tier[tIndx].requiredUnitLevel
          ]);
        }
        abilityData.push(abID,abName,unit.baseId,maxTier, "");
        if(options.images){
          abilityData.unshift(image);
        }
        if(options.tiers){
          tierData.forEach(tier => {
            abilityData.push(tier[0],tier[1],tier[2]);
          });
          for(let col=abilityData.length; col < headers.length; col++){
            abilityData.push("");
          }
        }
        if(append  && currentData.length > 0){
          if(abilityData.length !== writeData[0].length){
            let tempAbility = currentData.filter(ability => { return ability[idCol-1] === abID; })[0];
            for(let col=abilityData.length; col < currentData[0].length; col++){
              if(tempAbility === undefined){
                abilityData.push("");
              } else{
                abilityData.push(tempAbility[col]);
              }
            }
          }
        }
        writeData.push(abilityData.slice());
      });
      //=>Get Crew Skills
      unit.crew.forEach(crew => {
        crew.skillReference.forEach(skill => {
          abilityData = [];
          tierData = [];
          abID = skill.skillId;
          sIndx = skillMap[skill.skillId];
          aIndx = abilityMap[gameData["skill"][sIndx].abilityReference];
          abName = gameData["localization"][gameData["ability"][aIndx].nameKey];
          image = '=IMAGE("https://game-assets.swgoh.gg/' + gameData["ability"][aIndx].icon + '.png",1)';
          maxTier = gameData["skill"][sIndx].tier.length + 1;
          //=> Set tier data
          for(let i=0; i < (maxTier-1); i++ ){
            let tIndx = i;
            let rIndx = recipeMap[gameData["skill"][sIndx].tier[tIndx].recipeId];
            tierMat = "";
            gameData["recipe"][rIndx].ingredients.forEach(mat => {
              let mIndx = (materialMap.hasOwnProperty(mat.id)) ? materialMap[mat.id] : null;
              let matName = (mat.id === "SHIP_GRIND") ? gameData["localization"]["Shared_Currency_ShipGrind"] : gameData["localization"][gameData["material"][mIndx].nameKey];
              tierMat +=  matName + ":" + mat.maxQuantity + ",";
            });
            tierMat = tierMat.substring(0, tierMat.length-1);
            tierData.push([
              gameData["localization"][gameData["ability"][aIndx].tier[tIndx].descKey],
              tierMat,
              trCrewGearLv + ":" + gameData["skill"][sIndx].tier[tIndx].requiredUnitTier
            ]);
          }
          abilityData.push(abID,abName,unit.baseId,maxTier, crew.unitId);
          if(options.images){
            abilityData.unshift(image);
          }
          if(options.tiers){
            tierData.forEach(tier => {
              abilityData.push(tier[0],tier[1],tier[2]);
            });
            for(let col=abilityData.length; col < headers.length; col++){
              abilityData.push("");
            }
          }
          if(append  && currentData.length > 0){
            if(abilityData.length !== writeData[0].length){
              let tempAbility = currentData.filter(ability => { return ability[idCol-1] === abID; })[0];
              for(let col=abilityData.length; col < currentData[0].length; col++){
                if(tempAbility === undefined){
                  abilityData.push("");
                } else{
                  abilityData.push(tempAbility[col]);
                }
              }
            }
          }
        });
        writeData.push(abilityData.slice());
      });
    }
  });

  //Clear Sheet
  if(writeSheet.getRange("A1").getValue() !== ""){
    writeSheet.clearContents().clear()
  }
  //Write Sheet
  writeSheet.getRange(1,1, writeData.length, writeData[0].length).setValues(writeData);
  //Format sheet
  writeSheet.getRange(1,1,1,writeData[0].length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');  
  deleteUnusedCells(writeSheet);
  //writeSheet.autoResizeColumns(1,writeData[0].length);
  writeSheet.setFrozenRows(1);
  writeSheet.sort(idCol, true);
  writeSheet.sort(unitCol, true);

  return gameData; //Return data in case calling other functions that need it.
}


/**
 * Updates the gearData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} gameData - The currently retrieved gameData collections
 * @param {boolean} append - Flag to add new data changes instead of replacing everything
 * @returns {object} gameData - Returns the gameData to be used with other functions.
 */
function updateGearSheet(settingSheetName, gameData = null, append = false){
  const ws = new SheetMap(settingSheetName);
  //-->Create Sheet and designate write sheet
  var writeSheet
  var currentData = [];
  if(ws.gear_sheet === null){
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_gearSheet);
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_gearSheet);
  } else {
    writeSheet = ws.gear_sheet;
    currentData = (writeSheet.getRange("A1").getValue() === "") ? [] : writeSheet.getRange(1,1,writeSheet.getLastRow(),writeSheet.getLastColumn()).getValues();
  }
  //-->API settings
  let url = ws.api_url;
  let access = (ws.api_access === "") ? null : ws.api_access;
  let secret = (ws.api_secret === "") ? null : ws.api_secret;
  let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  if(gameData !== null){
    if(!gameData.hasOwnProperty("equipment")){ gameData['equipment'] = api.fetchData(['equipment']).equipment; }
    if(!gameData.hasOwnProperty("recipe")){ gameData['recipe'] = api.fetchData(['recipe']).recipe; }
    if(!gameData.hasOwnProperty("localization")){ gameData['localization'] = api.fetchLocalization(); }
  }else {
    gameData = api.fetchData(['equipment','recipe']);
    gameData['localization'] = api.fetchLocalization();
  }
  //==>Build Maps
  var equipmentMap = [];
  for(let i=0; i < gameData["equipment"].length; i++){
    equipmentMap[gameData["equipment"][i]["id"]] = i;
  }
  var recipeMap = [];
  for(let r=0; r < gameData["recipe"].length; r++){
    recipeMap[gameData["recipe"][r]["id"]] = r;
  }

  //Build Headers
  var writeData = [];
  let headers = [
      translate("gear ID", ws.api_language),
      translate("name",ws.api_language)
  ];
  for(let i=1; i < 6; i++){
    headers.push(
      translate("mat " + i + " id", ws.api_language),
      translate("mat " + i + " qty", ws.api_language)
    );
  }
  writeData.push(headers.slice());
  if(append && currentData.length > 0){
    for(let col=writeData[0].length; col < currentData[0].length; col++){
      writeData[0].push(currentData[0][col]);
    }
  }

  var gearData = [];
  gameData["equipment"].forEach(gear => {
    gearData = [];
    let recipeQty = 0;
    let gearName = gameData["localization"][gear.nameKey];
    gearData.push(gear.id,gearName);
    if(gear.recipeId !== ""){
      gameData["recipe"][recipeMap[gear.recipeId]].ingredients.forEach(function(mat){
        if( mat.id.toString().toLowerCase() !== "grind"){
          gearData.push(mat.id,mat.maxQuantity); 
          recipeQty += 1;
        }
      });
    }
    for(var i=0; i < (5 - recipeQty); i++){
      gearData.push("","");
    }
    if(append  && currentData.length > 0){
      if(gearData.length !== writeData[0].length){
        let tempGear = currentData.filter(g => { return g[1] === gearName; })[0];
        for(let col=gearData.length; col < currentData[0].length; col++){
          if(tempGear === undefined){
            gearData.push("");
          } else{
            gearData.push(tempGear[col]);
          }
        }
      }
    }  
    writeData.push(gearData.slice());
  });

  //Clear Sheet
  if(writeSheet.getRange("A1").getValue() !== ""){
    writeSheet.clearContents().clearFormats().clear();
  }
  //Write Sheet
  writeSheet.getRange(1,1, writeData.length, writeData[0].length).setValues(writeData);
  //Format sheet
  writeSheet.getRange(1,1,1,writeData[0].length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#cccccc');  
  deleteUnusedCells(writeSheet);
  //writeSheet.autoResizeColumns(1,writeData[0].length);
  writeSheet.getRangeList(['A2:A','C2:C',"E2:E",'G2:G','I2:I','K2:K']).setNumberFormat("000");
  writeSheet.getRangeList(['D2:D','F2:F',"H2:H",'J2:J','L2:L']).setNumberFormat("0");
  writeSheet.setFrozenRows(1);
  writeSheet.sort(1, true);

  return gameData; //Return data in case calling other functions that need it.
}


/**
 * Updates the relicData sheet
 * @param {string} settingSheetName - The name of the settings sheet
 * @param {object} gameData - The currently retrieved gameData collections
 * @param {boolean} append - Flag to add new data changes instead of replacing everything
 * @returns {object} gameData - Returns the gameData to be used with other functions.
 */
function updateRelicSheet(settingSheetName, gameData = null){
  const ws = new SheetMap(settingSheetName);
  //-->Create Sheet and designate write sheet
  var writeSheet
  var currentData = [];
  if(ws.relic_sheet === null){
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(ws.name_relicSheet);
    writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ws.name_relicSheet);
  } else {
    writeSheet = ws.relic_sheet;
    currentData = (writeSheet.getRange("A1").getValue() === "") ? [] : writeSheet.getRange(1,1,writeSheet.getLastRow(),writeSheet.getLastColumn()).getValues();
  }
  //-->API settings
    let url = ws.api_url;
    let access = (ws.api_access === "") ? null : ws.api_access;
    let secret = (ws.api_secret === "") ? null : ws.api_secret;
    let lang = (ws.api_language === "") ? "ENG_US" : getLocalizationLang(ws.api_language);
  const api = new SWGoHAPI.Comlink(url, access, secret, lang);
  //-->Get game data
  if(gameData !== null){
    if(!gameData.hasOwnProperty("material")){ gameData['material'] = api.fetchData(['material']).material; }
    if(!gameData.hasOwnProperty("recipe")){ gameData['recipe'] = api.fetchData(['recipe']).recipe; }
    if(!gameData.hasOwnProperty("localization")){ gameData['localization'] = api.fetchLocalization(); }
  }else {
    gameData = api.fetchData(['material','recipe']);
    gameData['localization'] = api.fetchLocalization();
  }

  //-->Get all relic mats
  var relicMats = [];
  gameData["material"].forEach(mat => {
    if(mat.type === 11 || mat.type === 12){
      relicMats.push([mat.id,gameData["localization"][mat.nameKey]]);
    }
  });
  relicMats.sort();
  relicMats.unshift(["GRIND",gameData["localization"]["Shared_Currency_Grind"]]);
  
  //Build Headers
  var writeData = [];
  let headerRow = [];
  for(let row=1; row < 3; row++){
    headerRow = [];
    let indx = (row === 1) ? 1 : 0;
    if(row === 1){
      headerRow.push("");
    }else{
      headerRow.push(translate("Tiers",ws.api_language));
    }
    for(let i=0; i < relicMats.length;i++){
      headerRow.push(relicMats[i][indx]);
    }
    writeData.push(headerRow.slice());
  }

  var relicRecipe = [];
  gameData["recipe"].forEach(recipe => {
    if(recipe["id"].includes("relic_promotion_recipe")){
      let relicLv = parseInt(recipe.id.substr(-2));
      let recipeMats = [];
      recipe.ingredients.forEach(mat => {
        recipeMats[mat.id] = mat.minQuantity;
      });
      relicRecipe[relicLv] = {
        id: "RELIC_" + relicLv,
        mats: recipeMats
      };
    }
  });

  const maxRows = Object.keys(relicRecipe).length;
  for(let row=1; row < maxRows + 1; row++){
    let tempData = [];
    tempData.push(relicRecipe[row].id);
    for(let col = 1; col < writeData[0].length; col++){
      if(relicRecipe[row].mats.hasOwnProperty( writeData[1][col])){
        tempData.push(relicRecipe[row].mats[ writeData[1][col] ]);
      } else {
        tempData.push("");
      }
    }
    writeData.push(tempData);
  }

  //Clear Sheet
  if(writeSheet.getRange("A2").getValue() !== ""){
    writeSheet.clearContents().clearFormats().clear();
    writeSheet.getBandings().forEach(banding => {
      banding.remove();
    });
  }
  //Write Sheet
  writeSheet.getRange(1,1, writeData.length, writeData[0].length).setValues(writeData).setBorder(true,true,true,true,true,true, 'black', SpreadsheetApp.BorderStyle.SOLID);
  //Format sheet
  writeSheet.getRange(1,1,1,writeData[0].length).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#45818e').setWrap(true);  
  writeSheet.getRange(2,1,1,writeData[0].length).setFontWeight('bold').setHorizontalAlignment('center');
  writeSheet.getRange(2,1, (writeData.length -1), writeData[0].length).applyRowBanding().setHeaderRowColor('#76a5af').setFirstRowColor('ffffff').setSecondRowColor('#ebeff1');
  deleteUnusedCells(writeSheet);
  //writeSheet.autoResizeColumns(1,writeData[0].length);
  writeSheet.setFrozenRows(2);

  return gameData; //Return data in case calling other functions that need it.
}
