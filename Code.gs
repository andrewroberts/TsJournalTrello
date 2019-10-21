/**
 * 'Sheet open' event handler
 */
 
function onOpen() {

  DocumentApp
    .getUi()
      .createMenu('TsJournalTrello')
        .addItem('TsJournalTrello', 'runTsJournalTrello')
        .addToUi()

} // onOpen()


function runTsJournalTrello() {
  
  // Get the current document ID and pass to the function to add 
  // a URL in the journal and hyperlink to the timesheet
  var DocId = DocumentApp.getActiveDocument().getId()
  var result = TsJounalTrello.tsJournalTrello(DocId)
  
  if (result === SUCCESS) {
    Logger.log('success!')
  } else {
    Logger.log('Sorry the script did not run correctly')
  }
}
