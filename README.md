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
Reload your Google Sheet, you will see the **Essential Sheet Menu** appear in the menu bar. Select **Help** then **Build Settings Sheet**. See STEP 3 in [Getting Started with Essential Sheet](#getting-started-with-essential-sheet) below for further instructions.

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
