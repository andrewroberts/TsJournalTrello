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
var SCRIPT_VERSION = "v0.1.dev_ajr"

//AJR Capital first letters indicate a service or constructor, so this would be tsJournalTrello()

//AJR We use something called JSDoc to automatically create docs from the code and these dictate 
// the format of the function headers, so along with my standard format:

// (I keep a template at the bottom of my config.gs - 
//
// https://github.com/andrewroberts/GAS-Framework/blob/master/Config.gs)

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

function TsJournalTrello() {
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
  //To do Check if the document link already exists.
  
//AJR Constants use capitals, e.g. TRELLO_CARD_NAME_. These I'd pull out into a separate
// config.gs file, and as this makes them global I'd add an underscore to avoid them being 
// visible outside the script should it be used as a library. Converting it to a library will
// be another exercise
  
  // I try and keep to Google's style guide: https://google.github.io/styleguide/jsguide.html
  
  //Config
  var trelloCardName = 'Example Feature 2 {Example Catagory 2}'
  var trellourl = 'https://trello.com/c/oNQk53mz/11-example-feature-1-examplecategory2'
  var timesheetrange = 'O2'
  var orgName = 'Acme Ltd'
  
  var CompanyCol = 1
  var TimesheetCol = 17
  
  //IDs
  var orgId = '1IwctVagVOgmlmGbJt0atVQPkyTp_ZKocoRkuQa4cEtU'
  //var timesheetId = '1wk717ZUEo50crwZr8fUXdutD1YproOzi4SughyjHwSU'
  var dociId = '12lKDk6A-P2-JGo4vhz1fE_33vC_MvxC1hbcUpfV5R2o' 
 
  //Open the Journal
  var journal = DocumentApp.openById(dociId)
  var body = journal.getBody()
  
//AJR - Use triple ===, and openById returns null if not found and you'll want to stop there, so:
//  if (journal === null) { 
//    Logger.log('Invalid Journal ID')
//    return
//  }
  
  if (journal == '') { 
    Logger.log('Invalid Journal ID')
  }

//AJR - If you are going to chain, chain them all

  //Add header 2 and hyperlink  in Document to trello task  
  body.appendParagraph(trelloCardName)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2).setLinkUrl(trellourl)
  
  //Get the timesheet link from the organisation spreadsheet
  var orgData = SpreadsheetApp.openById(orgId).getSheetByName('Organisations').getDataRange()
  
  var data = orgData.getValues()
  var header = data.shift()
  var timesheetUrl = ''

//AJR - Look at using Array.prototype.some() to search as it stops once it finds
// the matching value rather than looking through all of them like forEach does (well done
// for avoiding for(;;;) though, forEach is a lot quicker

    data.forEach(function(row) {
      
      //check if the company name matches the orgName
      var orgNameSearch = row[CompanyCol]
      if (orgNameSearch === orgName) {
        timesheetUrl = row[TimesheetCol]
        Logger.log(timesheetUrl)
        return timesheetUrl
      }
    })
    
    Logger.log(timesheetUrl)
 
//AJR - The timesheet links to the GDoc header and then the GDoc header links to the Trello card

    //Open timesheet and add the Trello Card and Hyperlink 
    //var timesheetTask = SpreadsheetApp.openById(timesheetId).getSheetByName('Timesheet').getRange(timesheetrange)
    if (timesheetUrl === '') {
      Logger.log('No Timesheet URL found in Org Spreadsheet')
    } else {
              
      var timesheetTask = SpreadsheetApp.openByUrl(timesheetUrl).getSheetByName('Timesheet').getRange(timesheetrange)
      timesheetTask.setFormula('=HYPERLINK("' + trellourl + '", " ' + trelloCardName + '")')
    }
}