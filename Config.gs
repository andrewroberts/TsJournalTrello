  // 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - 30Oct2019
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
var SCRIPT_VERSION = "v1.0"

var PRODUCTION_VERSION_ = true

// Log Library
// -----------

var DEBUG_LOG_LEVEL_ = PRODUCTION_VERSION_ ? BBLog.Level.INFO : BBLog.Level.FINER
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES_ = PRODUCTION_VERSION_ ? BBLog.DisplayFunctionNames.NO : BBLog.DisplayFunctionNames.NO
var DEBUG_LOG_SHEET_ID_ = '1WADNOsJwZdTqAV0a_N4yotTzJdgVIr4Zv7OoSP74BaM'

// Assert library
// --------------

var SEND_ERROR_EMAIL_ = PRODUCTION_VERSION_ ? true : false
var HANDLE_ERROR_ = Assert.HandleError.THROW

// Tests
// -----

var TEST_DOC_ID_ = '1Q-_uIu6S80Xm-Q_Do8N0i1BGtB6kNzIm4LR_8niWIfA' // ChCS - Journal

// Constants/Enums
// ===============

var START_URL_INDEX_ = 21
var END_URL_INDEX_ = 29
  
// Organisation Spreadsheet
var ORG_NAME_COLUMN_INDEX_ = 1
var ORG_TIMESHEET_COLUMN_INDEX_ = 18
var ORG_TRELLO_COLUMN_INDEX_ = 22
  
// Timesheet 
var TIMESHEET_NOTES_COLUMN_NUMBER_ = 15
var TIMESHEET_DATE_COLUMN_RANGE_ = "A1:A"

// Function Template
// -----------------

/**
 *
 *
 * @param {Object} 
 *
 * @return {Object}
 */
/* 
function functionTemplate() {

  Log_.functionEntryPoint()
  
  

} // functionTemplate() 
*/  