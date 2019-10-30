function test_init() {
  Log_ = BBLog.getLog({
    level:                DEBUG_LOG_LEVEL_, 
    displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
    sheetId:              DEBUG_LOG_SHEET_ID_, 
  })
}
function test_linkHeaders() {
  test_init()
  linkHeaders_(TEST_DOC_ID_)
  return
}
