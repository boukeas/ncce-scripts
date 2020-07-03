/*
function test() {
  date = new Date();
  var dd = date.getDate().toString().padStart(2,'0');
  var mm = (date.getMonth()+1).toString().padStart(2,'0');
  var yy = date.getFullYear().toString().slice(-2);
  dateString = dd + "-" + mm + "-" + yy;
  
  id = "1ufS_axhBGRqakuoVtWPH6r62vQCee2kjomarDEn2k10";
  
  result = lastChance(id, dateString);
  Logger.log(result);
}
*/

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function fixLastUpdated(date) {

  var indexSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index");
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Last updated").activate();
 
  var dateString;
  if (date) {
    dateString = date;
  } else {
    date = new Date();
    var dd = date.getDate().toString().padStart(2,'0');
    var mm = (date.getMonth()+1).toString().padStart(2,'0');
    var yy = date.getFullYear().toString().slice(-2);
    dateString = dd + "-" + mm + "-" + yy;
  }
    
  // clear log and create header row
  logSheet.clear();
  logSheet.appendRow(["Log message", "Folder", "File", "URL"]);
  logSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  var fixed = 0;
  var nbRows = indexSheet.getLastRow();
  range = indexSheet.getRange(2, 1, nbRows, 6);
  for (var row=1; row<=range.getNumRows(); row++) {
    if (!(isLink = range.getCell(row,3).getValue()) && (url = range.getCell(row,6).getValue())) 
      if (id = isDriveUrl(url)) 
        if (lastChance(id, dateString)) { 
          fixed++;
          logSheet.appendRow(["Updated", range.getCell(row,1).getValue(), range.getCell(row,2).getValue(), url]);
        } else {
          logSheet.appendRow(["Skipped", range.getCell(row,1).getValue(), range.getCell(row,2).getValue(), url]);
        }
  }
  return fixed;
}

//// auxillary

function lastChance(id, dateString) {  
  
  // Searches in the footer for "Last updated:" and 
  // *replaces* the following 8 characters with current date.
  // Caution: the ways this works is _not_ robust. The 8 chars
  // that follow the "Last updated:" string will be replaced
  // regardless of content. You shouldn't tamper with the template.
  
  var file = DriveApp.getFileById(id);
  if (file.getMimeType() == "application/vnd.google-apps.document") { 

    // this file is a document
    // look for the the text in the footer
    var doc = DocumentApp.openById(id);
    var parent = doc.getBody().getParent(); 
    for (var i = 0; i < parent.getNumChildren(); i += 1 ) {
      var childType = parent.getChild(i).getType();
      if (childType === DocumentApp.ElementType.BODY_SECTION || childType === DocumentApp.ElementType.HEADER_SECTION) continue; // not interested
      else if (childType === DocumentApp.ElementType.FOOTER_SECTION) {
        var footer = parent.getChild(i).asFooterSection();
        if (footer) {
          var start = footer.getText().toLowerCase().search("last updated:");
          if (start >= 0) {
            footer.editAsText().insertText(start+14, dateString);
            footer.editAsText().deleteText(start+22, start+29);
            return true;
          }
        }
      }
    }
    return false;

  } else if (file.getMimeType() == "application/vnd.google-apps.presentation") {

    // this file is a presentation
    // look for the text in the notes of the first slide
    var slides = SlidesApp.openById(id);
    var notesShape = slides.getSlides()[0].getNotesPage().getSpeakerNotesShape();
    var matches = notesShape.getText().find("Last updated: [0-9][0-9]-[0-9][0-9]-[0-9][0-9]");
    if (matches.length > 0) {
      // update existing date
      var txt = matches[0].getRange(14, matches[0].getLength());
      txt.clear(0, txt.getLength());
      txt.appendText(dateString);
      return true;
    } else {
      // insert "Last updated" string
      var txt = notesShape.getText().getRange(0,0);
      txt.appendText("Last updated: " + dateString + "\n\n");
      return true;
    }    
  }
}

// https://github.com/uxitten/polyfill/blob/master/string.polyfill.js
// https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/padStart
if (!String.prototype.padStart) {
    String.prototype.padStart = function padStart(targetLength,padString) {
        targetLength = targetLength>>0; //truncate if number or convert non-number to 0;
        padString = String((typeof padString !== 'undefined' ? padString : ' '));
        if (this.length > targetLength) {
            return String(this);
        }
        else {
            targetLength = targetLength-this.length;
            if (targetLength > padString.length) {
                padString += padString.repeat(targetLength/padString.length); //append to original to ensure we are longer than needed
            }
            return padString.slice(0,targetLength) + String(this);
        }
    };
}

function fixSelfrefs() {  
  
  var indexSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index");
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Self-reference").activate();

  // clear log and create header row
  logSheet.clear();
  logSheet.appendRow(["Log message", "Folder", "File", "Shortlink", "URL"]);
  logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  
  var fixed = 0;
  var nbRows = indexSheet.getLastRow();
  range = indexSheet.getRange(1, 1, nbRows, 6);  
  for (var row=2; row<=range.getNumRows(); row++) {
    if (!(isLink = range.getCell(row,3).getValue()) && (path = range.getCell(row,4).getValue())) 
      if (url = range.getCell(row,6).getValue()) 
        if (id = isDriveUrl(url)) 
          if (thisResource(id, path)) { 
            fixed++;
            logSheet.appendRow(["Updated", range.getCell(row,1).getValue(), range.getCell(row,2).getValue(), path, url]);
          } else {
            logSheet.appendRow(["Skipped", range.getCell(row,1).getValue(), range.getCell(row,2).getValue(), path, url]);
          }
  }
  return fixed;
}

//// auxillary

var thisResourceMatch = /This resource is available online at ncce.io\/([\A-Za-z0-9\-]+)/;

function thisResource(id, path) {
  
  var file = DriveApp.getFileById(id);
  if (file.getMimeType() == "application/vnd.google-apps.document") {
    
    // this file is a document
    // look for the the text in the body
    var doc = DocumentApp.openById(id);
    var editable = doc.getBody().editAsText();
    var match = thisResourceMatch.exec(editable.getText());
    if (match) {
      editable.setLinkUrl(match.index + 37, match.index + match[0].length - 1, "ncce.io/" + path);
      var replacementStart = match.index + match[0].length - match[1].length;
      editable.deleteText(replacementStart, replacementStart + match[1].length - 1);
      editable.insertText(replacementStart, path);
      return true;
    } else 
      
      // because this is a doc, but the text could not be found
      return false;
    
  } else if (file.getMimeType() == "application/vnd.google-apps.presentation") {

    // this file is a presentation
    // look for the text in the notes of the first slide
    var slides = SlidesApp.openById(id);
    var notesShape = slides.getSlides()[0].getNotesPage().getSpeakerNotesShape();
    var matches = notesShape.getText().find("This resource is available online at ncce.io\/([\A-Za-z0-9\-]+)");
    if (matches.length > 0) {
      var link = matches[0].getRange(37, matches[0].getLength());
      link.clear(8, link.getLength());
      link.appendText(path);
      link.getTextStyle().setLinkUrl("ncce.io/" + path);
      return true;
    } else {

      // look for the text in the elements of the first slide
      var firstSlideElements = slides.getSlides()[0].getPageElements();
      for (var i=0; i<firstSlideElements.length; i++) {
        var element = firstSlideElements[i];
        if (element.getPageElementType() == "SHAPE") {
          var matches = element.asShape().getText().find("This resource is available online at ncce.io\/([\A-Za-z0-9\-]+)");
          if (matches.length > 0) {
            var link = matches[0].getRange(37, matches[0].getLength());
            link.clear(8, link.getLength());
            link.appendText(path);
            link.getTextStyle().setLinkUrl("ncce.io/" + path);
            return true;
          }
        }
      }
      
      // because this is a presentation, but the text could not be found in the first slide
      // (neither in the notes, nor in the elements)
      return false;
    } 
  }
  
  // because this is neither a doc nor a presentation
  return false;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function updateSaveACopyDoc(id) {
  // for documents with different headers
  // https://stackoverflow.com/questions/31598708/google-apps-script-targeting-header-on-doc-with-different-first-page-headers
  
  var doc = DocumentApp.openById(id);
  var parent = doc.getBody().getParent(); 
  for (var i = 0; i < parent.getNumChildren(); i += 1 ) {
    var childType = parent.getChild(i).getType();
    if (childType === DocumentApp.ElementType.BODY_SECTION || childType === DocumentApp.ElementType.FOOTER_SECTION ) continue; // not interested
    else if (childType === DocumentApp.ElementType.HEADER_SECTION ) {
      var header = parent.getChild(i).asHeaderSection();
      if (header) {
        var start = header.getText().toLowerCase().search("save a copy");
        if (start >= 0) {
          header.editAsText().setLinkUrl(start, start+10, "https://docs.google.com/document/d/" + doc.getId() + "/copy");
          // [log]
          logger.appendRow(["located 'Save a copy link'", doc.getName(), doc.getId()]);
          break;
        }
      }
    }
  }
}

function updateSaveACopyPresentation(id) {
  var slides = SlidesApp.openById(id);
  var firstSlideElements = slides.getSlides()[0].getPageElements();
  for (var i=0; i<firstSlideElements.length; i++) {
    var element = firstSlideElements[i];
    if (element.getPageElementType() == "SHAPE") {
      var shape = element.asShape();
      var range = shape.getText().find("Save a copy");
      if (range.length > 0) {
        range[0].getTextStyle().setLinkUrl("https://docs.google.com/presentation/d/" + slides.getId() + "/copy");
        // [log]
        logger.appendRow(["located 'Save a copy link'", slides.getName(), slides.getId()]);
        break;
      }
    } else if (element.getPageElementType() == "TABLE") {
      var table = element.asTable();
      for (var row=0; row < table.getNumRows(); row++) {
        for (var col=0; col < table.getNumColumns(); col++) {
          var cell = table.getCell(row, col)
          var range = cell.getText().find("Save a copy");
          if (range.length > 0) {
            range[0].getTextStyle().setLinkUrl("https://docs.google.com/presentation/d/" + slides.getId() + "/copy");
            // [log]
            logger.appendRow(["located 'Save a copy link'", slides.getName(), slides.getId()]);
            break;
          }
        }
      }
    }
  }
}
