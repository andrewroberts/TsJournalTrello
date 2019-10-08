//Library - tsJournalTrello

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

//AJR - One feature I will use now is these, so we can practice tracking versions:
//
// We can call the latest code you pushed "v0.1". I've created a new branch for that.
// Note how this version that I'm not working on is a branch off v0.1 so I've added 
// the "dev_ajr" suffix and this is the name of the branch I'll do my daily pushes to 
// until it is ready to be released to either v0.2 or v1.0.

var SCRIPT_NAME = "TsJournalTrello"
var SCRIPT_VERSION = "v0.1.dev_dlt"

//AJR Capital first letters indicate a service or constructor, so this would be tsJournalTrello()

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
 * 2) The Timesheet Spreadsheet creating a hyperlink in the 'O2' 
 * column. The Timesheet Spreadsheet link is taken from the organisation name found in the 
 * organisation spreadsheet.
 *
 * @param {object} none
 *
 * @return {object} none
 */

function tsJournalTrello(DocId) {
  // Description: Function automates the process of copying the Trello 
  //              Card Name to the,
  //              1) Journal, creating a 'Heading 2' format text with a 
  //              Hyperlink to the Trello card.    
  //              2) The Timesheet Spreadsheet creating a hyperlink in the 'O2' 
  //              column. The Timesheet Spreadsheet link is taken from the organisation name found in the 
  //              organisation spreadsheet.
  // 
  // Author: Debbie Thomas
  // Date: 4th October 2019
  
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
        timesheetUrl = row[TIMESHEET_COL_]
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
 
      //Open the Journal
  var journal = DocumentApp.openById(DocId)
  var DocUrl = journal.getUrl()
  
  if (journal === null) { 
    Logger.log('Invalid Journal ID')
    return
  }
  
  var pars = journal.getBody().getParagraphs();  
  for(var i in pars) {
    var p = pars[i], h = p.getHeading(), d = DocumentApp.ParagraphHeading;
    if (h == DocumentApp.ParagraphHeading.HEADING2) {
      var TrelloCardTitle = p.getText()
      
      //Add the link to the URL to that paragraph
      //TODO: Change to the Url of the Trello Card
      //TODO: change to find the text after heading 1 and set the next line of text as heading 2
      p.setLinkUrl(trelloBoardUrl)
    }
  }


  // TODO: Get the lists of cards from the board ID and find the URL of the Card from TrelloCardTitle
   // var trelloBoardList = test_getBoardLists(trelloBoardId
   //Find the card name from the Board List
  
    //Open timesheet and add the Hyperlink to Journal
    if (timesheetUrl === null) {

      Logger.log('No Timesheet URL found in Org Spreadsheet')
      return
    } else {
              
      var timesheetTask = SpreadsheetApp.openByUrl(timesheetUrl).getSheetByName('Timesheet').getRange(TIMESHEET_RANGE_)
      timesheetTask.setFormula('=HYPERLINK("' + DocUrl + '", " ' + TrelloCardTitle + '")')
    }
return SUCCESS
}

function test_getBoardLists(trelloBoardId) {

  try {

    var trelloApp = new TrelloApp.App()    
    var boards = trelloApp.getBoardLists(trelloBoardId)
    Logger.log(boards)
  
  } catch (error) {

    if (error.name === 'AuthorizationError') {
    
      // This is a special error thrown by TrelloApp to indicate
      // that user authorization is required    
      showAuthorisationDialog()
      
    } else {
    
      throw error
    }
  }
  
  return
  
  // Private Functions
  // -----------------
  
  function showAuthorisationDialog() {
      
    var authorizationUrl = trelloApp.getAuthorizationUri()
    
    Dialog.show(
      'Opening authorization window...', 
        'Follow the instructions in this window, close ' + 
        'it and then try the action again. ' + 
        '<br/><br/>Look out for a warning that ' + 
        'your browser has blocked the authorisation pop-up from Trello. ' + 
        '<script>window.open("' + authorizationUrl + '")</script>',
      160)
      
  } // showAuthorisationDialog()
    
} // test_createCard()

function reset() {
  new TrelloApp.App().reset()
}
