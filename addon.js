/* What should the add-on do after it is installed */
function onInstall() {
  onOpen();
}

/* What should the add-on do when a document is opened */
function onOpen() {
  SpreadsheetApp.getUi()
  .createAddonMenu() // Add a new option in the Google Docs Add-ons Menu
  .addItem("Nouvelle demande", "createNewDemand")
  .addToUi();  // Run the showSidebar function when someone clicks the menu
}

/**
 * this is triggered by each value modification in every sheet
 * @param {[Object]} event (you can build it by hand for testing pupose)
 * milite-selection are discarded (no action)
 * add your specific handling inside
 * @return : none
 */
function onEdit(e) {
  // in case of multiple selection, e.value is undefined
  if( e.value === undefined) { return; } // multi-cell range

  //===== insert your specific handling below ====
  // column validation checked
  checkbox(e, "", 'B', [writeUserStamp, performNoteAction]);
}