/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////// MENUs
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// MANUAL JSON DOWNLOAD (MISSING FILE) AND REFRESH URLs MANUALLY (ALL OF THEM)
// .addItem --> The first parameter is the Name that will be displayed on the Menu, the Second parameter is the function this menu-object will call

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Manual JSON Download')
      .addItem('Company_1', 'Company_1_MissingJSON')
      .addSeparator()
      .addItem('Company_2', 'Company_2_MissingJSON')
      .addSeparator()
      .addItem('Company_3', 'Company_3_MissingJSON')
      .addSeparator()
      .addToUi();
   ui.createMenu('Refresh URLs')
      .addItem('Refresh all URLs Spreadsheets', 'Config_RefreshURLSpreadSheet')
      .addSeparator()
      .addToUi();
}
