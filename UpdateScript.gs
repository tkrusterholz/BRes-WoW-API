// Kazekiame  -  2023-10-27
// Call WoW Character API and retrieve data for spreadsheet purposes
// See the README at https://github.com/tkrusterholz/BRes-WoW-API/blob/main/README.md for important notes and instructions on how to use this!!!

function personalize() {
    //                                               ADD YOUR INFO HERE!!!
    // vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv      MANDATORY FIELDS      vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv
    PropertiesService.getScriptProperties().setProperty("realmName","Shadowsong")  // name of realm
    PropertiesService.getScriptProperties().setProperty("rangeOfNames","A1:A5")  // Cell range with character names
  
    // vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv      OPTIONAL FIELDS      vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv
    PropertiesService.getScriptProperties().setProperty("levelRange","")  // Cell range for character levels
    PropertiesService.getScriptProperties().setProperty("avgItemLevelRange","")  // Cell range for character average item levels
    PropertiesService.getScriptProperties().setProperty("eqptItemLevelRange","")  // Cell range for character equipped item levels
    PropertiesService.getScriptProperties().setProperty("genderRange","E1:E5")  // Cell range for character gender
    PropertiesService.getScriptProperties().setProperty("raceRange","F1:F5")  // Cell range for character race
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
         PropertiesService.getScriptProperties().getProperty('realmName').toLowerCase()+"/"+  // realm
         toon.toLowerCase()+"?namespace=profile-us&locale=en_US&access_token="+  // character
         PropertiesService.getScriptProperties().getProperty('tokenValue'),  // token
         {"muteHttpExceptions" : true}  // ignore HTTP errors
      ))
    }  
  }
  
  function update () {
    PropertiesService.getScriptProperties().deleteAllProperties()  // clean your room!
    personalize()  // gimme your info!
    PropertiesService.getScriptProperties().setProperty('tokenValue',BResWoWAPIToken.getToken())  // fetch API access token
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var names = sheet.getRange(PropertiesService.getScriptProperties().getProperty("rangeOfNames")).getValues()
    var masterArray = [[],[],[]]
  
    for (var rangeKey in PropertiesService.getScriptProperties().getProperties()) {
      if (rangeKey.indexOf("Range")>=0 && PropertiesService.getScriptProperties().getProperty(rangeKey) !== "") {
        masterArray[0].push(rangeKey)
        masterArray[1].push(PropertiesService.getScriptProperties().getProperty(rangeKey))
        masterArray[2].push([])
      }
    }
  
    for (var row in names) {
      for (var col in names[row]) {
  
        var toonData = getCharacter(names[row][col]) // Get character info from WoW API
  
        if (masterArray[0].includes("levelRange")) {
          if (toonData["code"] == 404) {var pasteData = ""} else {var pasteData = toonData["level"]}
          masterArray[2][masterArray[0].indexOf("levelRange")].push([pasteData]) }
        if (masterArray[0].includes("avgItemLevelRange")) {
          if (toonData["code"] == 404) {var pasteData = ""} else {var pasteData = toonData["average_item_level"]}
          masterArray[2][masterArray[0].indexOf("avgItemLevelRange")].push([pasteData]) }
        if (masterArray[0].includes("eqptItemLevelRange")) {
          if (toonData["code"] == 404) {var pasteData = ""} else {var pasteData = toonData["equipped_item_level"]}
          masterArray[2][masterArray[0].indexOf("eqptItemLevelRange")].push([pasteData]) }
        if (masterArray[0].includes("genderRange")) {
          if (toonData["code"] == 404) {var pasteData = ""} else {var pasteData = toonData["gender"]["name"]}
          masterArray[2][masterArray[0].indexOf("genderRange")].push([pasteData]) }
        if (masterArray[0].includes("raceRange")) {
          if (toonData["code"] == 404) {var pasteData = ""} else {var pasteData = toonData["race"]["name"]}
          masterArray[2][masterArray[0].indexOf("raceRange")].push([pasteData]) }                
      }
    }
  
    // Clear existing spreadsheet values and paste the data
    for (var pasteRange in masterArray[1]) {
      sheet.getRange(masterArray[1][pasteRange]).clearContent()
      sheet.getRange(masterArray[1][pasteRange]).setValues(masterArray[2][pasteRange])
    }
  }
