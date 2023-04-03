/**
 * Class for keeping up with all sheet names and fields required for this library.
 */
function SheetMap(settingSheetName) {
  var ws = SpreadsheetApp.getActiveSpreadsheet();

  //Sheets
  this.settings_sheet = ws.getSheetByName(settingSheetName);
  this.sheetNames = this.settings_sheet.getRange("J20:J31").getValues();
  this.name_guildSheet = this.sheetNames[0][0];
  this.name_playerSheet = this.sheetNames[1][0];
  this.name_rosterHeroSheet = this.sheetNames[2][0];
  this.name_rosterShipSheet = this.sheetNames[3][0];
  this.name_modSheet = this.sheetNames[4][0];
  this.name_datacronSheet = this.sheetNames[5][0];
  this.name_unitSheet = this.sheetNames[6][0];
  this.name_baseSheet = this.sheetNames[7][0];
  this.name_heroAbilitySheet = this.sheetNames[8][0];
  this.name_shipAbilitySheet = this.sheetNames[9][0];
  this.name_gearSheet = this.sheetNames[10][0];
  this.name_relicSheet = this.sheetNames[11][0];
  this.guild_sheet = ws.getSheetByName(this.name_guildSheet);
  this.player_sheet = ws.getSheetByName(this.name_playerSheet);
  this.unit_sheet = ws.getSheetByName(this.name_unitSheet);
  this.base_sheet = ws.getSheetByName(this.name_baseSheet);
  this.hero_ability_sheet = ws.getSheetByName(this.name_heroAbilitySheet);
  this.ship_ability_sheet = ws.getSheetByName(this.name_shipAbilitySheet);
  this.gear_sheet = ws.getSheetByName(this.name_gearSheet);
  this.relic_sheet = ws.getSheetByName(this.name_relicSheet);
  this.roster_hero_sheet = ws.getSheetByName(this.name_rosterHeroSheet);
  this.roster_ship_sheet = ws.getSheetByName(this.name_rosterShipSheet);
  this.mod_sheet = ws.getSheetByName(this.name_modSheet);
  this.datacron_sheet = ws.getSheetByName(this.name_datacronSheet);

  //Named Ranges - API Setup
  this.loaded_guilds = this.settings_sheet.getRange("B4:F28").getValues();
  this.loaded_players = this.settings_sheet.getRange("B32:F56").getValues();
  this.input_guildID = this.settings_sheet.getRange("I9").getValue();
  this.input_playerID = this.settings_sheet.getRange("I10").getValue();
  this.api_url = this.settings_sheet.getRange("I13").getValue();
  this.api_access = this.settings_sheet.getRange("I14").getValue();
  this.api_secret = this.settings_sheet.getRange("I15").getValue();
  this.api_language = this.settings_sheet.getRange("I16").getValue();

  //Named Ranges - Create Sheets
  //this.create_guild = this.settings_sheet.getRange("I20").getValue();
  //this.create_player = this.settings_sheet.getRange("I21").getValue();
  this.create_heroRoster = this.settings_sheet.getRange("I22").getValue();
  this.create_shipRoster = this.settings_sheet.getRange("I23").getValue();
  this.create_mod = this.settings_sheet.getRange("I24").getValue();
  this.create_datacron = this.settings_sheet.getRange("I25").getValue();
  this.create_units = this.settings_sheet.getRange("I26").getValue();
  this.create_heroBase = this.settings_sheet.getRange("I27").getValue();
  this.create_heroAbility = this.settings_sheet.getRange("I28").getValue();
  this.create_shipAbility = this.settings_sheet.getRange("I29").getValue();
  this.create_gear = this.settings_sheet.getRange("I30").getValue();
  this.create_relic = this.settings_sheet.getRange("I31").getValue();

  //Named Ranges - Guild Options
  this.option_guild_event = this.settings_sheet.getRange("I35").getValue();

  //Named Ranges - Member Options
  this.option_member_gp = this.settings_sheet.getRange("I37").getValue();
  this.option_member_time = this.settings_sheet.getRange("I38").getValue();
  this.option_member_gac = this.settings_sheet.getRange("I39").getValue();
  this.option_member_tickets = this.settings_sheet.getRange("I40").getValue();
  this.option_member_raids = this.settings_sheet.getRange("I41").getValue();

  //Named Ranges - Hero Options
  this.option_hero_stats = this.settings_sheet.getRange("I43").getValue();
  this.option_hero_abilities = this.settings_sheet.getRange("I44").getValue();
  this.option_hero_tiers = this.settings_sheet.getRange("I45").getValue();

  //Named Ranges - Mod Options
  this.option_mod_rolls = this.settings_sheet.getRange("I47").getValue();
  this.option_mod_stats = this.settings_sheet.getRange("I48").getValue();

  //Named Ranges - Ship Options
  this.option_ship_stats = this.settings_sheet.getRange("I50").getValue();
  this.option_ship_tiers = this.settings_sheet.getRange("I51").getValue();

  //Named Ranges - Unit Options
  this.option_unit_images = this.settings_sheet.getRange("I53").getValue();

  //Named Ranges - Unit Base Options
  this.option_base_stats = this.settings_sheet.getRange("I55").getValue();

  //Named Ranges - Ability Options
  this.option_ability_images = this.settings_sheet.getRange("I57").getValue();
  this.option_ability_tiers = this.settings_sheet.getRange("I58").getValue();

}
