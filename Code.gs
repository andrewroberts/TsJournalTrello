// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// Code review all files - TODO
// JSHint review (see files) - TODO
// Unit Tests - TODO
// System Test (Dev) - TODO
// System Test (Prod) - TODO

// Config.gs
// =========
//
// Dev: AndrewRoberts.net
//
// All the constants and configuration settings

// Configuration
// =============

var SCRIPT_NAME = "TsJournalTrello"
var SCRIPT_VERSION = "v0.1.dev_dlt"

var PRODUCTION_VERSION_ = false

// Log Library
// -----------

var DEBUG_LOG_LEVEL_ = PRODUCTION_VERSION_ ? Log.Level.INFO : Log.Level.FINER
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES_ = PRODUCTION_VERSION_ ? Log.DisplayFunctionNames.NO : Log.DisplayFunctionNames.YES

// Assert library
// --------------

var SEND_ERROR_EMAIL_ = PRODUCTION_VERSION_ ? true : false
var HANDLE_ERROR_ = Assert.HandleError.THROW
var ADMIN_EMAIL_ADDRESS_ = ''

// Tests
// -----

var TEST_SHEET_ID_ = ''

// Constants/Enums
// ===============



// Function Template
// -----------------

/**
 *
 *
 * @param {object} 
 *
 * @return {object}
 */
function tsJournalTrello(DocId) {

  //Open the organisation spreadsheet
  var orgData = SpreadsheetApp.openById(ORG_ID_)
      .getSheetByName('Organisations')
      .getDataRange()
  
  var data = orgData.getValues()
  var header = data.shift()
  var timesheetUrl = null
  var trelloBoardUrl = null

  // Get the organisation name from the journal name, text before '- journal' 
  var str = ' - Journal'
  var CHAR_TO_END_ = str.length
  var orgName = DocumentApp.getActiveDocument().getName().slice(0, -CHAR_TO_END_)
    
  //Loop through the rows of data
  data.some(function(row)  {
      
    //check if the company name matches the orgName and if so, get the timesheet and
    //Trello Board Url
    var orgNameSearch = row[ORG_NAME_COL_]
      
    if (orgNameSearch === orgName) {
      timesheetUrl = row[ORG_TS_COL_]
      trelloBoardUrl = row[TRELLO_COL_]
      return
      } 
    })    
    
    Logger.log('Timesheet URL: ' + timesheetUrl)
    Logger.log('Trello Board URL: ' + trelloBoardUrl)
    
    //Check the Trello Board Exists  
    if (trelloBoardUrl === null) {
      Logger.log('No Trello Board URL found in Org Spreadsheet')
      return
    } else {
      
      //Get the Trello Board Id from the wedsite json return
      try {
   
        var result = UrlFetchApp.fetch(trelloBoardUrl + '.json', {muteHttpExceptions:true})
        var response = result.getContentText()
        var trelloBoardData = JSON.parse(response) 
        var trelloBoardId = trelloBoardData.id   
      } catch (error) {
    
        Log.warning(url + ' not accessible: ' + error.message)
        return
      }
    }
 
  //Open the Journal, loop through the paragraphs until the Heading 1 is found
  //The next paragraph is the Trello Card Name, change the style to Heading 2 and add a 
  //Hyperlink to the Trello Card 
  var journal = DocumentApp.openById(DocId)
  var DocUrl = journal.getUrl()
 
  if (journal === null) { 
    Logger.log('Invalid Journal ID')
    return
  }
  
  var pars = journal.getBody().getParagraphs();  
  for(var i in pars) {
    var p = pars[i], h = p.getHeading(), d = DocumentApp.ParagraphHeading;
    if (h == DocumentApp.ParagraphHeading.HEADING1) {
      // Set the next paragraph to heading 2 and add the Trello Card Title URL link
      var np = pars[Number(i) + 1]
      var nptext = np.getText()
      Logger.log(nptext)
      np.setHeading(DocumentApp.ParagraphHeading.HEADING2)

      //Store the text as the Trello Card Title
      var trelloCardTitle = nptext
      
      //Get the Url from the card title, add the link to the journal
      var trelloCardUrl = getTrelloCardUrl(trelloBoardId, trelloCardTitle)
      
      if (trelloCardUrl === null) {
      Logger.log('Trello card Not found in trello board')
      }
      
      // Add the Hyperlink to the journal
      np.setLinkUrl(trelloCardUrl)
      
     }
  }
  
  //Open the Timesheet spreadsheet
  var tsData = SpreadsheetApp.openByUrl(timesheetUrl)
    .getSheetByName('Timesheet')
    .getRange("A1:A").getValues()
  //Get the last row with data in the Date column, use this to add the Trello Card and 
  //Journal URL link to the timesheet
  var lastTsRow = tsData.filter(String).length 
  Logger.log(lastTsRow)
  
  //Open timesheet and add the Hyperlink to Journal
  if (timesheetUrl === null) {

    Logger.log('No Timesheet URL found in Org Spreadsheet')
    return
    
  } else {
              
    var timesheetTask = SpreadsheetApp.openByUrl(timesheetUrl)
      .getSheetByName('Timesheet')
      .getRange(lastTsRow, TIMESHEET_COL_)
    timesheetTask.setFormula('=HYPERLINK("' + DocUrl + '", " ' + trelloCardTitle + '")')
    
  }
    
return SUCCESS

}

function getTrelloCardUrl(trelloBoardId, trelloCardTitle) {
  
  //Get the JSON response of the Trello Board
  var API_KEY = PropertiesService.getScriptProperties().getProperty("API_KEY")
  var TOKEN = PropertiesService.getScriptProperties().getProperty("TOKEN")
  var url = "https://api.trello.com/1/boards/" + trelloBoardId + "/cards/?fields=name,url&key=" + API_KEY + "&token=" + TOKEN
  var response = UrlFetchApp.fetch(url).getContentText()
  var cardUrl = null
  
  var obj = JSON.parse(response)
        
  //Find the card title in the list of board cards   
  for (x in obj) {
    if (obj[x].name === trelloCardTitle) {
      cardUrl = obj[x].url
      Logger.log(cardUrl)
      return
    } 
  }
 
  if (cardUrl === null) {
    Logger.log('Card Name not found in list of Trello Boards')
    return
  }
  
}