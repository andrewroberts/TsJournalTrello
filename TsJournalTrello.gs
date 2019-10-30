// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 30Oct2019
/* jshint asi: true */
/* jshint esversion: 6 */

(function() {"use strict"})()

// TsJournalTrello.gs
// ==================
//
// External interface to this script - all of the event handlers.
//
// This files contains all of the event handlers, plus miscellaneous functions 
// not worthy of their own files yet
//
// The filename is prepended with _API as the Github chrome extension won't 
// push a file with the same name as the project.

var Log_
var Properties_

// Public event handlers
// ---------------------
//
// All external event handlers need to be top-level function calls; they can't 
// be part of an object, and to ensure they are all processed similarily 
// for things like logging and error handling, they all go through 
// errorHandler_(). These can be called from custom menus, web apps, 
// triggers, etc
// 
// The main functionality of a call is in a function with the same name but 
// post-fixed with an underscore (to indicate it is private to the script)
//
// For debug, rather than production builds, lower level functions are exposed
// in the menu

var EVENT_HANDLERS_ = {

//                           Name                            onError Message                          Main Functionality
//                           ----                            ---------------                          ------------------

  linkHeaders:                 ['linkHeaders()',                 'linkHeaders Failed',                     linkHeaders_],
}

function linkHeaders(args) {return eventHandler_(EVENT_HANDLERS_.linkHeaders, args)}

// Private Functions
// =================

// General
// -------

/**
 * All external function calls should call this to ensure standard 
 * processing - logging, errors, etc - is always done.
 *
 * @param {Array} config:
 *   [0] {Function} prefunction
 *   [1] {String} eventName
 *   [2] {String} onErrorMessage
 *   [3] {Function} mainFunction
 *
 * @param {Object}   args       The argument passed to the top-level event handler
 */

function eventHandler_(config, args) {

  Properties_ = args.properties
  
  try {

    var userEmail = Session.getActiveUser().getEmail()
    var logSheetId = Properties_.getProperty("LOG_SHEET_ID")
    
    Log_ = BBLog.getLog({
      level:                DEBUG_LOG_LEVEL_, 
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
      sheetId:              logSheetId, 
    })
    
    Log_.info('Handling ' + config[0] + ' from ' + (userEmail || 'unknown email') + ' (' + SCRIPT_NAME + ' ' + SCRIPT_VERSION + ')')
    
    // Call the main function
    return config[2](args.id)
    
  } catch (error) {

    var handleError = Assert.HandleError.DISPLAY_FULL
  
    if (!PRODUCTION_VERSION_) {
      handleError = Assert.HandleError.THROW
    }
  
    var assertConfig = {
      error:          error,
      userMessage:    config[1],
      log:            Log_,
      handleError:    handleError, 
      sendErrorEmail: SEND_ERROR_EMAIL_, 
      emailAddress:   Session.getEffectiveUser().getEmail(),
      scriptName:     SCRIPT_NAME,
      scriptVersion:  SCRIPT_VERSION,
    }
  
    Assert.handleError(assertConfig)
  }
    
} // eventHandler_()

// Private event handlers
// ----------------------

/**
 * Link the timesheet entry to the journal, and the journal to the Trello card
 *
 * @param {string} docId
 */
 
function linkHeaders_(docId) {

  Log_.functionEntryPoint()
  
  var journal = DocumentApp.openById(docId)
  
  if (journal === null) {   
    throw new Error('Invalid Journal ID ' + docId)
  }
    
  var orgId = Properties_.getProperty('ORG_SHEET_ID')
  
  var orgData = SpreadsheetApp.openById(orgId)
    .getSheetByName('Organisations')
    .getDataRange()
  
  var data = orgData.getValues()
  var timesheetUrl = null
  var trelloBoardUrl = null
  var trelloBoardId = null

  //Get the organisation name from the journal name, text before '- journal' 
  var str = ' - Journal'
  var CHAR_TO_END_ = str.length
  var orgName = journal.getName().slice(0, -CHAR_TO_END_)
  
  //Look for this org
  var orgFound = false
    
  orgFound = data.some(function(row)  {
  
    var orgNameSearch = row[ORG_NAME_COLUMN_INDEX_]
      
    if (orgNameSearch === orgName) {    
      timesheetUrl = row[ORG_TIMESHEET_COLUMN_INDEX_]
      trelloBoardUrl = row[ORG_TRELLO_COLUMN_INDEX_]
      return true      
    }
  })   
    
  if (!orgFound) {
    throw new Error('Organisation Name: ' + orgName + ' not found in Org Spreadsheet')
  } 
  
  Log_.info('Organisation Name: ' + orgName + ' found in Org Spreadsheet')
    
  if (trelloBoardUrl === null) {
    throw new Error('No Trello Board URL found in Org Spreadsheet')
  } 
  
  Log_.info('Timesheet URL: ' + timesheetUrl)
    
  trelloBoardId = getTrelloBoardId_(trelloBoardUrl)
     
  var [trelloCardTitle, npBookmarkId] = processJournal_(journal, trelloBoardId)
  
  if (trelloCardTitle === null && npBookmarkId === null) {
    Log_.warning('Trello title already complete')
    return
  } 
    
  addLinkToTimesheet_(timesheetUrl, npBookmarkId, journal, trelloCardTitle) 
  Log_.info('Header links complete!')
                
} // linkHeaders_() 

/**
 * Loop through the journal paragraphs until the Heading 1 is found. The next 
 * paragraph is the Trello Card Name. Check it hasnt already been processed.
 * If it hasnt, change the style to Heading 2 and add a Hyperlink to the Trello Card 
 *
 * @param {GDocument} journal
 * @param {string} trelloBoardId
 */

function processJournal_(journal, trelloBoardId) {

  Log_.functionEntryPoint()
            
  //Create an array of heading 1 paragraph indexes and use this to find the 
  //last one
  
  var paragraphs = journal.getBody().getParagraphs()
  var heading1Paragraphs = []
    
  paragraphs.forEach(function(paragraph, paragraphIndex) {
    if (paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
      heading1Paragraphs.push(paragraphIndex)
    }
  })  
  
  var lastHeading1Paragraph = heading1Paragraphs[heading1Paragraphs.length - 1]
  var trelloCardParagraph = paragraphs[lastHeading1Paragraph + 1]
  var trelloCardTitle = trelloCardParagraph.getText()
  
  if (trelloCardTitle === '') {
    throw new Error('No Trello card text')
  }
  
  //Check if the Trello link has already been added by seeing if it is heading 2   
  if (trelloCardParagraph.getHeading() === DocumentApp.ParagraphHeading.HEADING2) {
    Log_.warning('Trello Card Title already processed')
    return [null, null]
  }
  
  var trelloCardUrl = getTrelloCardUrl_(trelloBoardId, trelloCardTitle)  
  var sectionPos = journal.newPosition(trelloCardParagraph, 0);
  var npBookmark = journal.addBookmark(sectionPos)
  
  //Set the next paragraph to heading 2 and add the Trello Card Title URL link
  trelloCardParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING2)
  trelloCardParagraph.setLinkUrl(trelloCardUrl)
  
  var processJournalReturn = [trelloCardTitle, npBookmark.getId()]
  return processJournalReturn 
       
} // processJournal_()

/**
 * Get the last row with data in the Date column, use this to add the Trello Card and 
 * Journal URL link to the timesheet
 *
 * @param {string} timesheetUrl
 * @param {string} npBookmarkId
 * @param {GDocument} journal
 * @param {string} trelloCardTitle
 */

function addLinkToTimesheet_(timesheetUrl, npBookmarkId, journal, trelloCardTitle) {
  
  Log_.functionEntryPoint()
  
  var timeSheet = SpreadsheetApp.openByUrl(timesheetUrl).getSheetByName('Timesheet')
  var timesheetDates = timeSheet.getRange(TIMESHEET_DATE_COLUMN_RANGE_).getValues()
  
  var lastTimesheetRowNumber = null
  
  timesheetDates.some(function(date, rowIndex) {
    if (date[0] === '') {
      lastTimesheetRowNumber = rowIndex
      return true
    }
  })

  if (lastTimesheetRowNumber === null) {
    throw new Error('Failed to find the last row in the timesheet')
  }

  var timesheetTaskRange = timeSheet.getRange(lastTimesheetRowNumber, TIMESHEET_NOTES_COLUMN_NUMBER_)  
  var cellValue = timesheetTaskRange.getValue()
  
  if (cellValue !== '') {
    throw new Error('Data found in Timesheet Notes, clear the Task/Notes cell and try again')
  }
    
  var docUrl = journal.getUrl()      
  var formula = '=HYPERLINK("' + docUrl + '#bookmark=' + npBookmarkId + '", "' + trelloCardTitle + '")'
  timesheetTaskRange.setFormula(formula)
  Log_.fine('formula: ' + formula)
    
} // addLinkToTimesheet_()

/**
 * Get the Trello card URL from the Trello API
 *
 * @param {string} trelloBoardId
 * @param {string} trelloCardTitle
 *
 * @return {string} trelloCardUrl
 */

function getTrelloCardUrl_(trelloBoardId, trelloCardTitle) {

  Log_.functionEntryPoint()
  
  var apiKey = Properties_.getProperty("API_KEY")
  var token = Properties_.getProperty("TOKEN")
  var url = "https://api.trello.com/1/boards/" + trelloBoardId + "/cards/?fields=name,url&key=" + apiKey + "&token=" + token
  var response = UrlFetchApp.fetch(url).getContentText()
  var trelloCardUrl = null  
  var obj = JSON.parse(response)
        
  //Find the card title in the list of board cards   
  for (var key in obj) {  
    if (obj[key].name === trelloCardTitle) {    
      trelloCardUrl = obj[key].url
      Log_.info('Trellocard URL: ' + trelloCardUrl)
      break  
    }     
  }
  
  if (trelloCardUrl === null) {      
    throw new Error('Card Name ' + trelloCardTitle + ' not found in list of Trello Boards')
  }
  
  return trelloCardUrl 
  
} // getTrelloCardUrl_()

/** 
 * Get the Trello Board Id from the wedsite json return  
 *
 * @param {string} trelloBoardUrl
 *
 * @return {string} trelloBoardId
 */ 

function getTrelloBoardId_(trelloBoardUrl) {
 
  var apiKey = Properties_.getProperty("API_KEY")
  var token = Properties_.getProperty("TOKEN")
  
  //The Short Link from the trello board url is elements of the URL
  var trelloBoardShortLink = trelloBoardUrl.slice(START_URL_INDEX_, END_URL_INDEX_)        
  Log_.info('Trello Board Short Link ' + trelloBoardShortLink)
  
  var result = UrlFetchApp.fetch("https://api.trello.com/1/boards/" + trelloBoardShortLink + "?fields=name,url,shortId&key=" + apiKey + "&token=" + token)
  var trelloBoardId = JSON.parse(result).id
        
  if (trelloBoardId === undefined) {
    throw new Error('Trello Board Id ' + trelloBoardUrl + ' not found')
  }
  
  Log_.info('Trello Card Id: ' + trelloBoardId)
  return trelloBoardId
  
} // getTrelloBoardId_()
