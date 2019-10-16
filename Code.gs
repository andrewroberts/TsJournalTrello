//AJR - As further practice in doing version control with Apps Script, I've pulled
//the master branch into this new stand-alone script, and we can push/pull changes 
//through our two copies

//AJR - This is the standard boilerplate I use to start my projects:
//
// https://github.com/andrewroberts/GAS-Framework.
//
// Just take a look through it now, but a good exercise will be to bring this
// code into it eventually. I'll walk you through that process.

//AJR - Tip: I remembered when running this in the debugger first time: 
// On the first run after you finish the auth flow, the function will run all
// the way through, even if you don't want it. So I usually add a "return" 
// after the function name until the auth flow is finished.

var SCRIPT_NAME = "TsJournalTrello"
var SCRIPT_VERSION = "v0.1.dev_dlt"


//AJR We use something called JSDoc to automatically create docs from the code and these dictate 
// the format of the function headers, so along with my standard format:

// (I keep a template at the bottom of my config.gs - 
//
// https://github.com/andrewroberts/GAS-Framework/blob/master/Config.gs)
  // I try and keep to Google's style guide: https://google.github.io/styleguide/jsguide.html

/**
 * Function automates the process of copying the Trello Card Name to the,
 * 1) Journal, creating a 'Heading 2' format text with a 
 * Hyperlink to the Trello card.    
 * 2) The Timesheet Spreadsheet creating a hyperlink to the journal in the Task/Notes 
 * column. 
 *
 * @param {object} none
 *
 * @return {object} none
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
  //The next paragraph is the Trello Card Name, change the name to Heading 2 and add a 
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
      var TrelloCardTitle = nptext
      
      //Get the Url from the card title, add the link to the journal
      // TODO: Get the lists of cards from the board ID and find the URL of the Card from TrelloCardTitle
      var trelloCardUrl = getTrelloCardUrl(TrelloCardTitle)
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
  

   //Find the card name from the Board List
  
    //Open timesheet and add the Hyperlink to Journal
    if (timesheetUrl === null) {

      Logger.log('No Timesheet URL found in Org Spreadsheet')
      return
    } else {
              
      var timesheetTask = SpreadsheetApp.openByUrl(timesheetUrl)
      .getSheetByName('Timesheet')
      .getRange(lastTsRow, TIMESHEET_COL_)
      timesheetTask.setFormula('=HYPERLINK("' + DocUrl + '", " ' + TrelloCardTitle + '")')
    }
return SUCCESS
}

function getTrelloCardUrl(TrelloCardTitle) {
  
  var API_KEY = PropertiesService.getScriptProperties().getProperty("API_KEY")
  var TOKEN = PropertiesService.getScriptProperties().getProperty("TOKEN")
  var url = "https://api.trello.com/1/boards/5d91c934ef30f98eee00a4d2/cards/?fields=name,url&key=" + API_KEY + "&token=" + TOKEN
  var response = UrlFetchApp.fetch(url).getContentText()
  
  var cardUrl = null
  var obj = JSON.parse(response)
        
  for (x in obj) {
    if (obj[x].name === TrelloCardTitle) {
      cardUrl = obj[x].url
      Logger.log(cardUrl)
      return cardUrl
    } 
  }

  return
  
}
