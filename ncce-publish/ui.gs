	// globals
var unitInfo;

///

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('NCCE')
      .addItem('Publish', 'showIndexSidebar')
      .addToUi();
}

function showIndexSidebar() {
  var htmlOutput = HtmlService
  .createHtmlOutputFromFile('sidebar-index')
  .setTitle('NCCE Publish')
  .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function checkId(id) {
  try {
    file = DriveApp.getFolderById(id);
    return file.getName();
  } catch(error) {
    return null;
  }
}

function startIndexing(params) {
  index_init(params);
  SpreadsheetApp.getUi().alert("Indexing complete!");
  return unitInfo.tags;
}

function startChecking(shortcmCheck) {
  var nbIssues = check(shortcmCheck);
  if (nbIssues) {
    SpreadsheetApp.getUi().alert("Check complete!\n"+ nbIssues + " issues require your attention.");
  } else {
    SpreadsheetApp.getUi().alert("Check complete!\nYou can proceed with shortlink creation.");
  }
  return nbIssues == 0;
}

function startCreating(unitTags) {
  var nbRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index").getLastRow();
  if (nbRows) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Begin shortlink creation?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      create_init(unitTags);
      SpreadsheetApp.getUi().alert("Shortlink creation complete!");
    }
  } else 
    SpreadsheetApp.getUi().alert("Index is empty!");
}

function startReplacing(dryRun) {
  var nbRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index").getLastRow();
  if (nbRows) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Begin replacing links with shortlinks?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      replace_init(dryRun);
      SpreadsheetApp.getUi().alert("Links replaced.");
    }
  } else 
    SpreadsheetApp.getUi().alert("Index is empty!");
}

function startFixingUpdated(date) {
  var nbRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index").getLastRow();
  if (nbRows) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Fix "Last updated" dates in documents?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      var nbFixed = fixLastUpdated(date);
      SpreadsheetApp.getUi().alert('Fixed "Last updated" dates in ' + nbFixed + " documents.");
    }
  } else 
    SpreadsheetApp.getUi().alert("Index is empty!");
}

function startFixingSelfrefs() {
  var nbRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index").getLastRow();
  if (nbRows) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Fix self references in documents and slides?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      var nbFixed = fixSelfrefs();
      SpreadsheetApp.getUi().alert('Fixed self references in ' + nbFixed + " documents.");
    }
  } else 
    SpreadsheetApp.getUi().alert("Index is empty!");
}

// function clear() {
//  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().clearContents();
// }
