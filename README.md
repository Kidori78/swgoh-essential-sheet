**Translation:**
[中文(简体)](/readme/readme_chs_cn.md)
 | English
 | [Français](/readme/readme_fre_fr.md)
 | [Deutsch](/readme/readme_ger_de.md)
 | [Indonesia](/readme/readme_ind_id.md)
 | [Italiano](/readme/readme_ita_it.md)
 | [日本語](/readme/readme_jpn_jp.md)
 | [Português (Brasil)](/readme/reamde_por_br.md)
 | [Русский](/readme/readme_rus_ru.md)
 | [Español](/readme/readme_spa_xm.md)
 | [Türkçe](/readme/readme_tur_tr.md)
 
# SWGoH Essential Sheet
The SWGoH Essential Sheet is a Google Sheets Library that utilizes Comlink to retrieve data directly from the Star Wars Galaxy of Heroes mobile game and store it directly in the sheet. It provides game data, guild data, and player data including guild event tracking, daily raid tickets, base stats at each gear level for characters, gear recipes and relic recipes, and much more. It can be translated into all the same languages the game itself can and allows for renaming the tabs. It is provided as a library so it can be inserted into any project you currently have or the official sheet can be copied. 

This repo is used to store the library files and information for the SWGoH Essential Sheet. To learn more about Comlink visit the Official Github for the projectat [swgoh-utils/swgoh-comlink](https://github.com/swgoh-utils/swgoh-comlink). For help setting up your own version of Comlink on a hosting service see [swgoh-utils/deploy-swgoh-comlink](https://github.com/swgoh-utils/deploy-swgoh-comlink). For running it locally on your device and storing the data in your own GIthub repo rather than using a hosting service see [comlink-for-github](https://github.com/Kidori78/swgoh-comlink-for-github).

Get additional help and support for this sheet on my discord [Kidori's SWGoH Tools](https://discord.gg/Z3fWRWa77W).

<br />

## Adding to an existing spreadsheet
### STEP 1
Copy the Code.gs file above and paste it into your current Google Sheet by going to _Extensions_ then _Apps Script_, you can also copy it straight from the Essential Sheet. Adjust the property `sheetName` within it to whatever you would like it to be. Adjust the property `language` to your preferred language.\
Language Options:
* **English**  |  **中文(简体)**  |  **中國傳統的**  |  **Français**  |  **Deutsch**
* **Indonesia**  |  **Italiano**  |  **日本語**  |  **한국인**  |  **Português (Brasil)**
* **Русский**  |  **Español**  |  **แบบไทย**  |  **Türkçe**

**Example**\
_PropertiesService.getUserProperties().setProperties({`sheetName`: **"essentialSettings"**, `language`: **"English"**});_

### STEP 2
Add the Essential Sheet Library to Apps Script by clicking the + next to _Libraries_. Select the latest version and give it the identifier **EssentialSheet**.\
`1Ygs0Z2bkJLSt9in4QGaDFNEq1VzhAfKuumZClQuWY45c-9Kx7je87kih`

### STEP 3
Add the Google Sheets API Service to Apps Script by clicking the + next to _Services_. Select v4 and then give it the identifier **SheetsAPI**.

### STEP 4
Reload your Google Sheet, you will see the **Essential Sheet Menu** appear in the menu bar. Select **Help** then **Build Settings Sheet**. See [STEP 3 in Getting Started with Essential Sheet](#step-3-1) below for further instructions.

<br />

## Getting Started with Essential Sheet
### STEP 1
In the menu bar go to **Extensions** then **Apps Script**. Within **Code.gs** adjust the property `sheetName` to whatever you would like it to be. Adjust the property `language` to your preferred language.\
Language Options:
* **English**  |  **中文(简体)**  |  **中國傳統的**  |  **Français**  |  **Deutsch**
* **Indonesia**  |  **Italiano**  |  **日本語**  |  **한국인**  |  **Português (Brasil)**
* **Русский**  |  **Español**  |  **แบบไทย**  |  **Türkçe**

**Example**\
_PropertiesService.getUserProperties().setProperties({`sheetName`: **"essentialSettings"**, `language`: **"English"**});_

### STEP 2
If you changed the language reload the sheet, otherwise go to the menu bar at the top and select **Essential Sheet Menu**, then select **Help**, then select **Build Settings Sheet**.

### STEP 3
On the Settings sheet enter the web address where comlink can be found for Host URL, any access key and secret key your Comlink has set.

### STEP 4 
In the Available Sheets table adjust the name of each sheet to what you want it to be. Then click the checkbox for each sheet tab you would like to have created when loading user data and game data.

### STEP 5
In the Sheet Options table select any additional details you would like to include in each sheet tab.

### STEP 6
Enter an id next to GET GUILD and/or GET PLAYER, then go to the menu bar and select **Essential Sheet Menu**, then select **Load data**, then select either **Load guild data** or **Load player data**.\
_NOTES:_
* _The more GP a guild has, the longer it will take to load the data. Some guilds may take up to the full 6 minutes to load. If any guild time outs then you must load Mod Data by itself in a separate execution._
* _GET PLAYER can take multiple ids by separating them with a , and it cannot have any blank spaces in it._

### STEP 7
If you want to remove a guild or player, check the remove box next to them on the settings sheet and go to the menu bar and select **Essential Sheet Menu**, then select **Remove data**, and then select either **Remove guild data** or **Remove player data** depending on the option you want to do.\
_NOTE: Removing more than one guild at a time can make the sheet or the web browser freeze due to how much data changes on it and may require reloading the Sheet._

<br />

# Advanced Uses
The Essential Sheet Library was built in a way so you can utilize any of its functions or included Libraries to help expand your own Google Sheet projects.

## Included Libraries
The following libraries are included in the Essential Sheet Library and you can access them in the ways listed below. To learn what methods and functions they have please visit the appropriate documentation for them.

### SWGoH Comlink API
This library is a client wrapper for Comlink which handles all of the calls and responses to and from it. You would use this to grab your own data collections and user data responses.\
**Library ID:** SWGoHAPI\
**Access Methods:** EssentialLibraryID.ComlinkLibraryID.methods(arguments).   E.G. _EssentialSheet.SWGoHAPI.fetchGuilds(guildID)_\
**Documentation:** [https://github.com/Kidori78/swgoh-comlink-wrapper-for-gas](https://github.com/Kidori78/swgoh-comlink-wrapper-for-gas)

### SWGoH Stat Calculator
This library is used to build all of the player stats and galactic power.\
**Library ID:** StatsAPI\
**Access Methods:** EssentialLibraryID.StatsLibraryID.methods(arguments).   E.G. _EssentialSheet.StatsAPI.calcCharStats(charData)_\
**Documentation:** https://github.com/Kidori78/swgoh-stat-calc/tree/GAS-v2](https://github.com/Kidori78/swgoh-stat-calc/tree/GAS-v2)

## Essential Sheet Functions
All functions of the Essential Sheet Library can be accessed by calling the identifier first, EssentialSheet.loadGameData(settingSheetName). Below is a list of available functions that can be called.

### loadGameData(settingSheetName)
Loads all the selected sheet tabs with the game data.\
**Parameter:**\
`settingSheetName` Required\
The name of the settings sheet. Required to get sheet tab names and ranges.

### updateUnitDataSheet(settingSheetName, gameData, append)
Loads the unitData sheet with each character and ship in the game. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\
`append` Boolean | Optional\
Retains any added columns to the sheet when rebuilding it.

### updateBaseDataSheet(settingSheetName, gameData, append)
Loads the heroBaseData sheet with each character at each gear level. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\
`append` Boolean | Optional\
Retains any added columns to the sheet when rebuilding it.

### updateHeroAbilitySheet(settingSheetName, gameData, append)
Loads the heroAbilityData sheet with each character ability. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\
`append` Boolean | Optional\
Retains any added columns to the sheet when rebuilding it.

### updateShipAbilitySheet(settingSheetName, gameData, append)
Loads the shipAbilityData sheet with each ship ability. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\
`append` Boolean | Optional\
Retains any added columns to the sheet when rebuilding it.

### updateGearSheet(settingSheetName, gameData, append)
Loads the gearData sheet with all in-game gear. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\
`append` Boolean | Optional\
Retains any added columns to the sheet when rebuilding it.

### updateRelicSheet(settingSheetName, gameData, append)
Loads the relicData sheet with all relic tier recipes. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\
`append` Boolean | Optional\
Retains any added columns to the sheet when rebuilding it.

### loadGuildData(settingSheetName)
Loads all selected guild data into the selected sheet tabs.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\

### loadPlayerData(settingSheetName)
Loads all selected player data into the selected sheet tabs.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\

### updateGuildDataSheet(settingSheetName, gameData)
Loads user data into the guildData sheet. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\

### updatePlayerDataSheet(settingSheetName, gameData)
Loads user data into the playerData sheet. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\

### updateHeroRosterDataSheet(settingSheetName, gameData)
Loads user data into the heroRosterData sheet. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\

### updateShipRosterDataSheet(settingSheetName, gameData)
Loads user data into the shipRosterData sheet. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\

### updateModDataSheet(settingSheetName, gameData)
Loads user data into the modData sheet. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\

### updateDatacronDataSheet(settingSheetName, gameData)
Loads user data into the datacronData sheet. This function returns the gameData object.\
**Parameter:**\
`settingSheetName` String | Required\
The name of the settings sheet. Required to get sheet tab names and ranges.\
`gameData` Object | Optional\
The gameData object can be passed to this function to reduce api calls to keep grabbing it.\

### deleteUnusedCells(sheet, addCol)
Deletes all unused columns and rows on a specified sheet.\
**Parameter:**\
`sheet` Object | Required\
The sheet object to delete cells on.\
`addCol` Integer | Optional\
The number of empty columns to leave on the sheet

### getLocalizationLang(language)
Converts the string language into the accepted games iso format for that language.\
**Parameter:**\
`language` String | Required\

### translate(message, targetLanguage)
Translates a string to the target language using Googles translation services. There is a 300 translation limit per day.\
**Parameter:**\
`message` String | Required\
The message you want to translate.\
`targetLanguage` String | Required\
The language string or iso format to convert it into.

### sortSheets(targetSheet,columns)
Sort a sheet by the specified columns in ascending order.\
**Parameter:**\
`targetSheet` Object | Required\
The sheet object to sort.\
`columns` Array of Integers | Required\
An array that contains each column you wish to sort. Column 1 starts at 0.

### deleteRows(targetSheet)
Deletes all unused rows on a sheet.\
**Parameter:**\
`targetSheet` Object | Required\
The sheet object.\

### deleteEntries(targetSheet,players,allyCol,cleanup)
Clears the specified sheet of all data not specified in the columns.\
**Parameter:**\
`targetSheet` Object | Required\
The sheet object.\
`players` Array | Required\
The array of ids or other unique identifiers to get rid of a sheet. If cleanup is enabled this indicates the data to keep.\
`allyCol` Integer | Required\
The column the ally code for players can be found in. Can also indicate the column the unique identifier is found in.\
`cleanup` Boolean | Optional\
Specifies that the players array indicates the only values to keep on the sheet.

### appendData(targetSheet, data, raw)
Uses the Google Sheets API to add data to the end of a sheet. Allows for adding lots of data quickly and with a lower chance of it freezing the Google Sheet.\
**Parameter:**\
`targetSheet` Object | Required\
The sheet object to write to.\
`data` Array | Required\
The array of data to add to the sheet.\
`raw` Boolean | Optional\
Indicates if the data should be added as raw or user_entered. Raw means the data is added as strings only and no sheet calculations or guesses will be made for each cell to see what format it is. When false this will add data as if the user entered it on the sheet and Google attempts to determine if it is a function, number, date, etc. Default is true to always use raw.

### getOmicronArea(mode)
Returns the area of the game where the omicron will be in effect.\
**Parameter:**\
`mode` Integer | Required\
The mode id.\
