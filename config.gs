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

var SCRIPT_NAME = "tsJournalTrello"
var SCRIPT_VERSION = "v0.2"

var PRODUCTION_VERSION_ = false

// Log Library
// -----------

var DEBUG_LOG_LEVEL_ = PRODUCTION_VERSION_ ? BBLog.Level.INFO : BBLog.Level.FINER
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES_ = PRODUCTION_VERSION_ ? BBLog.DisplayFunctionNames.NO : BBLog.DisplayFunctionNames.NO

// Assert library
// --------------

var SEND_ERROR_EMAIL_ = PRODUCTION_VERSION_ ? true : false
var HANDLE_ERROR_ = Assert.HandleError.THROW
var ADMIN_EMAIL_ADDRESS_ = 'debbiethomas35@gmail.com'

// Constants/Enums
// ===============

var SUCCESS = true

//Organisation Spreadsheet
var ORG_NAME_COL_ = 1
var ORG_ID_ = '1IwctVagVOgmlmGbJt0atVQPkyTp_ZKocoRkuQa4cEtU'
var ORG_TS_COL_ = 17
  
// Timesheet 
var TIMESHEET_COL_ = 15
var TRELLO_COL_ = 21
var DATE_COL_ = "A2:A"

var START_URL_INDEX_ = 21
var END_URL_INDEX_ = 29
  
// Function Template
// -----------------

/**
 *
 *
 * @param {Object} 
 *
 * @return {Object}
 */
 
function functionTemplate() {

  Log_.functionEntryPoint()
  
  

} // functionTemplate() 
  
