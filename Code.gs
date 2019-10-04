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
  
  if (journal == '') {
    Logger.log('Invalid Journal ID')
  }

  //Add header 2 and hyperlink  in Document to trello task  
  body.appendParagraph(trelloCardName)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2).setLinkUrl(trellourl)
  
  //Get the timesheet link from the organisation spreadsheet
  var orgData = SpreadsheetApp.openById(orgId).getSheetByName('Organisations').getDataRange()
  
  var data = orgData.getValues()
  var header = data.shift()
  var timesheetUrl = ''
  
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
  
    //Open timesheet and add the Trello Card and Hyperlink 
    //var timesheetTask = SpreadsheetApp.openById(timesheetId).getSheetByName('Timesheet').getRange(timesheetrange)
    if (timesheetUrl === '') {
      Logger.log('No Timesheet URL found in Org Spreadsheet')
    } else {
              
      var timesheetTask = SpreadsheetApp.openByUrl(timesheetUrl).getSheetByName('Timesheet').getRange(timesheetrange)
      timesheetTask.setFormula('=HYPERLINK("' + trellourl + '", " ' + trelloCardName + '")')
    }
}
