// Kazekiame  -  2023-10-27
// Call WoW Character API and retrieve data for spreadsheet purposes
// See the README at https://github.com/tkrusterholz/BRes-WoW-API/blob/main/README.md for important notes and instructions on how to use this!!!

function personalize() {
  // vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv  ADD YOUR INFO HERE!!!  vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv
  PropertiesService.getUserProperties().setProperty("realmName","Shadowsong") // Realm name
  PropertiesService.getUserProperties().setProperty("nameRange","A2:A27") // Cell range with character names
  PropertiesService.getUserProperties().setProperty("levelRange","B2:B27") // Cell range with character levels
  PropertiesService.getUserProperties().setProperty("ilvlRange","C2:C27") // Cell range with character item levels
}

function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('WoW API').addItem('Update Character Info','update').addToUi()
}

function getCharacter(toon) { 
  if (toon == "") {
    return {"code":404}
  } else {
    // makethecall.gif
    return JSON.parse(UrlFetchApp.fetch("https://us.api.blizzard.com/profile/wow/character/"+
       PropertiesService.getUserProperties().getProperty('realmName').toLowerCase()+"/"+  // realm
       toon.toLowerCase()+"?namespace=profile-us&locale=en_US&access_token="+  // character
       PropertiesService.getUserProperties().getProperty('tokenValue'),  // token
       {"muteHttpExceptions" : true}  // ignore HTTP errors
    ))
  }  
}

function update () {
  personalize()  // gimme your infos
  PropertiesService.getUserProperties().setProperty('tokenValue',BResWoWAPIToken.getToken())  // fetch API access token
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var names = sheet.getRange(PropertiesService.getUserProperties().getProperty("nameRange")).getValues()
  var levelRange = PropertiesService.getUserProperties().getProperty("levelRange")
  var ilvlRange = PropertiesService.getUserProperties().getProperty("ilvlRange")
  var levelData = []
  var ilvlData = []

  for (var row in names) {
    for (var col in names[row]) {

      var toonData = getCharacter(names[row][col]) // Get individual character info from WoW API
      if (toonData["code"] == 404) {
        levelData.push([[]])
        ilvlData.push([[]])
      } else {
        levelData.push([toonData["level"].toFixed(0)]) // Insert character level into 1x1 array and append to level array
        ilvlData.push([toonData["average_item_level"].toFixed(0)]) // Insert character item level into 1x1 array and append to ilvl array
      }
    }
  }

  // Clear existing spreadsheet values and paste the data
  sheet.getRange(levelRange).clearContent()
  sheet.getRange(levelRange).setValues(levelData)
  sheet.getRange(ilvlRange).clearContent()
  sheet.getRange(ilvlRange).setValues(ilvlData)
}
