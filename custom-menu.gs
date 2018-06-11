/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////// MENUs
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// MANUAL JSON DOWNLOAD (MISSING FILE) AND REFRESH URLs MANUALLY (ALL OF THEM)
// .addItem --> The first parameter is the Name that will be displayed on the Menu, the Second parameter is the function this menu-object will call

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Manual JSON Download')
      .addItem('MYKITA_DE', 'MYKITA_DE_MissingJSON')
      .addSeparator()
      .addItem('MYKITA_UK', 'MYKITA_UK_MissingJSON')
      .addSeparator()
      .addItem('MYKITA_US', 'MYKITA_US_MissingJSON')
      .addSeparator()
      .addItem('GGB_EN', 'GGB_EN_MissingJSON')
      .addSeparator()
      .addItem('GGB_DE', 'GGB_DE_MissingJSON')
      .addSeparator()
      .addItem('GGB_FR', 'GGB_FR_MissingJSON')
      .addSeparator()
      .addToUi();
   ui.createMenu('Refresh URLs')
      .addItem('Refresh all URLs Spreadsheets', 'Config_RefreshURLSpreadSheet')
      .addSeparator()
      .addToUi();
}
