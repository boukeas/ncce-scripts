function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('NCCE Clone')
      .addItem('Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var htmlOutput = HtmlService
  .createHtmlOutputFromFile('sidebar')
  .setTitle('NCCE Clone')
  .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function init(params) {
  sourceId = params.sourceId;
  parentId = params.parentId;
  dryRun = params.dryRun;
  updateShortlinksFlag = params.linkUpdate;
  prefix = params.prefix;
  
  switch(params.access) {
    case "0":
      defaultAccess = DriveApp.Access.ANYONE;
      break;
    case "1":
      defaultAccess = DriveApp.Access.ANYONE_WITH_LINK;
      break;
    case "2":
      defaultAccess = DriveApp.Access.DOMAIN;
      break;
    case "3":
      defaultAccess = DriveApp.Access.DOMAIN_WITH_LINK;
      break;
    case "4":
      defaultAccess = DriveApp.Access.PRIVATE;
      break;
  }
  switch(params.permission) {
    case "0":
      defaultPermission = DriveApp.Permission.VIEW;
      break;
    case "1":
      defaultPermission = DriveApp.Permission.COMMENT;
      break;
    case "2":
      defaultPermission = DriveApp.Permission.EDIT;
      break;
  }
  
  clone_init();
  
  // Here's the feature you requested Andy! 
  SpreadsheetApp.getUi().alert("Cloning complete!")
}

function clear() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().clearContents();
}

function checkId(id) {
  try {
    file = DriveApp.getFolderById(id);
    return file.getName();
  } catch(error) {
    return null;
  }
}

////// sources / references
//
// sidebar
// ** https://stackoverflow.com/questions/54713239/how-to-send-inputs-from-google-spreadsheet-sidebar-into-sheet-script-function
// https://developers.google.com/apps-script/guides/dialogs
// https://developers.google.com/apps-script/guides/html/
// https://github.com/tomcam/gassidebar
// https://subscription.packtpub.com/book/web_development/9781785882517/2/ch02lvl1sec22/creating-a-sidebar
// https://www.benlcollins.com/apps-script-examples/

////// client-server communication
// https://developers.google.com/apps-script/guides/html/communication
// http://ramblings.mcpher.com/Home/excelquirks/gassnips/progresshtml

////// folder picker (forget about this, too complicated)
// https://ctrlq.org/code/20039-google-picker-with-apps-script
