// Kazekiame  -  2023-10-27
// Call WoW Character API and retrieve data for spreadsheet purposes
// See the README at https://github.com/tkrusterholz/BRes-WoW-API/blob/main/README.md for important notes and instructions on how to use this!!!
// Updated 2025-02-18 for new Blizzard API auth headers

function personalize() {
    //                                               ADD YOUR INFO HERE!!!
    // vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv      MANDATORY FIELDS      vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv
    PropertiesService.getScriptProperties().setProperty("realmName","Shadowsong")  // name of realm
    PropertiesService.getScriptProperties().setProperty("rangeOfNames","A2:A51")  // Cell range with character names
  
    // vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv      OPTIONAL FIELDS      vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv vvv
    PropertiesService.getScriptProperties().setProperty("levelRange","B2:B51")  // Cell range for character levels
    PropertiesService.getScriptProperties().setProperty("avgItemLevelRange","C2:C51")  // Cell range for character average item levels
    PropertiesService.getScriptProperties().setProperty("eqptItemLevelRange","")  // Cell range for character equipped item levels
    PropertiesService.getScriptProperties().setProperty("genderRange","")  // Cell range for character gender
    PropertiesService.getScriptProperties().setProperty("raceRange","D2:D51")  // Cell range for character race
  }
  
  function onOpen(e) {
    SpreadsheetApp.getUi().createMenu('WoW API').addItem('Update Character Info','update').addToUi()
  }
  
  function getCharacter(toon) { 
    var headers = {
      "Battlenet-Namespace" : "profile-us",
      "Authorization" : "Bearer "+PropertiesService.getScriptProperties().getProperty('tokenValue')}
    const options = {headers:headers, muteHttpExceptions:true}  // ignore HTTP errors
    if (toon == "") {
      return {"code":404}
    } else {
      // makethecall.gif
      var response = UrlFetchApp.fetch("https://us.api.blizzard.com/profile/wow/character/"+
         PropertiesService.getScriptProperties().getProperty('realmName').toLowerCase()+"/"+  // realm
         toon.toLowerCase()+"?locale=en_US",options)  // character
      return JSON.parse(response)
    }
  }
  
  function update () {
    //PropertiesService.getScriptProperties().setProperty('realmName',"shadowsong")
    PropertiesService.getScriptProperties().deleteAllProperties()  // clean your room!
    personalize()  // gimme your info!
    PropertiesService.getScriptProperties().setProperty('tokenValue',BResWoWAPIToken.getToken())  // fetch API access token
    Logger.log(PropertiesService.getScriptProperties().getProperty('tokenValue'))
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

        Logger.log(toonData)
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
