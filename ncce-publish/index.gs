// Creates an index spreadsheet in your home Google Drive for a folder that follows NCCE naming conventions

// [TODOs]
// - Consider using Range.getValues, rather than getValue for each cell

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// The ids of files or folders included in this array will *not* be cloned.
var excludedIds = [];
// Folders whose names start with any of the prefixes included in this list will *not* be cloned.
var excludedPrefixes = ["[Ignore]", "[Residual]"];

var indexSheet;
var documents;
var short2long;
var long2short;

var indexSheet;
var documents;
var short2long;
var long2short;
 
/// Run function init()
function index_init(params) {

  var folderId = params.folderId;
  var unit = params.unit;
    
  indexSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index").activate();

  // clear index and create header row
  indexSheet.clear();
  indexSheet.appendRow(["Folder", "File", "Link text", "Shortlink", "Tags", "URL"]);
  indexSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  
  // clear other sheets as well
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Check").clear();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create").clear();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Replace").clear();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Self-reference").clear();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Last updated").clear();
  
  // look for folder to be indexed
  if (!folderId) return;
  folder = DriveApp.getFolderById(folderId);
  foldername = folder.getName();
  
  // get unit information (and an initial set of unit-related tags) from foldername
  unitInfo = parseRootname(foldername);
  if (unitInfo) {
    unitInfo["tags"] = ["Resource repository", "KS" + unitInfo.stage];
    if (unitInfo.strandOpt) unitInfo.tags.push(unitInfo.strandOpt);
    if (unitInfo.year) unitInfo.tags.push("Year " + unitInfo.year);
    unitInfo.tags.push(unitInfo.unitName);
  } else {
    unitInfo = {"tags": []};
  }

  if (unit) {
    unitInfo["unit"] = unit;
    unitInfo.tags.push(unit);
  }
    
  documents = [];
  // start indexing files in folders
  index(folder);  
    
  indexSheet.getRange(indexSheet.getLastRow(), 1, 1,7).setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  // after indexing, go through all documents
  // and locate "modlinks" (modifiable links), i.e. links followed by brackets
  documents.forEach(function (id) {
    var doc = DocumentApp.openById(id);
    locateModlinks(doc.getBody().editAsText()).forEach(function (modlink) {
      if (!modlink.path) modlink["path"] = "";
      indexSheet.appendRow([
        DriveApp.getFileById(id).getParents().next().getName(), doc.getName(), 
        modlink.text, modlink.path, " ", cleanDriveUrl(modlink.url), isDriveUrl(cleanDriveUrl(modlink.url))]);
      // indexSheet.getRange(indexSheet.getLastRow(), 1, 1, 6).setBackground("#f2f2f2");
    });
  });
}

function index(folder) {

  var foldername = folder.getName();
  var lesson;
  var link;
  var url;
  var id;
  
  // for all files in source  
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var filename = file.getName();
    
    // check for file exclusions
    if (excludedIds.indexOf(file.getId()) >= 0) continue;
    if (prefixMembership(file.getName(), excludedPrefixes)) continue;
    
    // use filename and foldername to suggest shortlink path and tags
    var suggestions = suggest(filename, foldername);
    
    // add a row for the file to the index
    indexSheet.appendRow([
      foldername, filename, "",
      suggestions.path, suggestions.tags.join(" | "), 
      url = cleanDriveUrl(file.getUrl()),
      id = file.getId()]);
                
    // if the file is a document, add it to the list of documents (to check for modlinks)
    if (file.getMimeType() == "application/vnd.google-apps.document") documents.push(id);
  }
    
  // for all folders in source
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
     
    if (excludedIds.indexOf(subFolder.getId()) >= 0) continue;
    if (prefixMembership(subFolder.getName(), excludedPrefixes)) continue;

    // index folder contents (recursive call)
    index(subFolder);
  }    
}

function check(shortcmCheck) {
  
  var nbIssues = 0;
  var indexSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index");
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Check").activate();
  
  // clear log and create header row
  logSheet.clear();
  logSheet.appendRow(["Message", "Index Line", "Additional information"]);
  logSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  
  var nbRows = indexSheet.getLastRow();
  range = indexSheet.getRange(1, 1, nbRows, 7);
  range.setBackground(null);
  
  var link;
  // mapping from shortlink paths to rows
  var path2rows = {};
  var rows;
  // mapping from long URL to shortlink path 
  var long2short = {};
  // mapping from long URL to file row 
  var long2row = {};
  // mapping from id to rows
  var id2rows = {};
  
  ////
  var urlMap = {};
  var pathMap = {};
  var idMap = {};
  
  // go through all file rows  
  row = 2;
  while (row <= nbRows && !range.getCell(row,3).getValue()) {
    if (url = range.getCell(row,6).getValue()) {  // is there a file URL?
      
      url = cleanDriveUrl(url);
      path = range.getCell(row,4).getValue();
      id = range.getCell(row,7).getValue();
      
      ////
      if (urlRecord = urlMap[url]) {
        urlRecord.rows.push(row);
      } else {
        urlRecord = (urlMap[url] = {"rows": [row], "paths": {}});
      }
      if (path && !urlRecord.paths[path]) urlRecord.paths[path] = true;

      ////
      if (path) {
        if (pathRecord = pathMap[path]) {
          pathRecord.rows.push(row);
        } else {
          pathRecord = (pathMap[path] = {"rows": [row], "urls": {}});
        }
        pathRecord.urls[url] = true;
      }

      ////
      if (id) {
        if (idRecord = idMap[id]) {
          idRecord.rows.push(row);
        } else {
          idRecord = (idMap[id] = {"rows": [row], "paths": {}, "urls": {}});
        }
        idRecord.urls[url] = true;
        if (path && !idRecord.paths[path]) idRecord.paths[path] = true;
      }

      /*
      // add the row to the mapping between id and rows
      if (id = range.getCell(row,7).getValue()) if (rows = id2rows[id]) rows.push(row); else id2rows[id] = [row];
      
      if (path = range.getCell(row,4).getValue()) { // is there a file shortlink?        
        // record the mapping from the longURL to the suggested shortlink path and the current row 
        long2short[url] = path;
        long2row[url] = row;
        // add the row to the mapping between path and rows
        if (rows = path2rows[path]) rows.push(row); else path2rows[path] = [row];
      } else {
        // record the mapping from the longURL to the current row:
        // you may need access to it later, in case a shortlink path is provided through a modlink
        long2row[url] = row;
      }
      */
    } else { 
      // error: there is no file URL, the user must have deleted it!
      //// [tmp] nbIssues++;
      // this shouldn't have happened: files are indexed with a URL.
      //// [tmp] logSheet.appendRow(["Error: missing file URL!", row, "Missing URL"]);
      
      /// [tmp]
      path = range.getCell(row,4).getValue();
      id = range.getCell(row,7).getValue();
      
      ////
      if (path) {
        if (pathRecord = pathMap[path]) {
          pathRecord.rows.push(row);
        } else {
          pathRecord = (pathMap[path] = {"rows": [row], "urls": {}});
        }
      }

      ////
      if (id) {
        if (idRecord = idMap[id]) {
          idRecord.rows.push(row);
        } else {
          idRecord = (idMap[id] = {"rows": [row], "paths": {}, "urls": {}});
        }
        logSheet.appendRow(["Check: missing file URL!", row, "Missing URL", id]);
        if (path && !idRecord.paths[path]) idRecord.paths[path] = true;
      }
    }
    row++;
  }
   
  /*
  // display and mark duplicate file shortlinks  
  for (var path in path2rows) {
    rows = path2rows[path]; 
    if (rows.length > 1) {
      nbIssues++;
      logSheet.appendRow(["Error: duplicate shortlink path", rows.join(", "), path]);
      for (var i=0; i<rows.length; i++) {
        range.getCell(rows[i], 4).setBackground("#ffd0d0");
      }
    }
  }
  */
  
  modlinksRow = row;
  while (row <= nbRows) {    
    if (url = range.getCell(row,6).getValue()) {  // is there a modlink URL?
      
      url = cleanDriveUrl(url);
      path = range.getCell(row,4).getValue();
      id = range.getCell(row,7).getValue();
      
      ////
      if (urlRecord = urlMap[url]) {
        urlRecord.rows.push(row);
      } else {
        urlRecord = (urlMap[url] = {"rows": [row], "paths": {}});
      }
      if (path && !urlRecord.paths[path]) urlRecord.paths[path] = true;      
      
      ////
      if (path) {
        if (pathRecord = pathMap[path]) {
          pathRecord.rows.push(row);
        } else {
          pathRecord = (pathMap[path] = {"rows": [row], "urls": {}});
        }
        pathRecord.urls[url] = true;
      }

      ////
      if (id) {
        if (idRecord = idMap[id]) {
          idRecord.rows.push(row);
        } else {
          idRecord = (idMap[id] = {"rows": [row], "paths": {}, "urls": {}});
        }
        idRecord.urls[url] = true;
        if (path && !idRecord.paths[path]) idRecord.paths[path] = true;
      }

      /*
      // add the row to the mapping between id and rows
      if (id = range.getCell(row,7).getValue()) if (rows = id2rows[id]) rows.push(row); else id2rows[id] = [row];

      
      if (path = range.getCell(row,4).getValue()) { // is there a modlink shortlink for the modlink URL?   

        // add the row to the mapping between path and rows
        if (rows = path2rows[path]) rows.push(row); else path2rows[path] = [row];
        
        if (mappedPath = long2short[url]) { // is there also a file shortlink for the modlink URL?
          if (path == mappedPath) { // does the modlink shortlink coincide with the file shortlink?
            // there is a modlink shortlink for the modlink URL
            // there is also a file shortlink for the modlink URL
            // the shortlink paths coincide
            // action: all good! (the shortlink can be created)
            // range.getCell(row, 4).setBackground("#f2f2f2");
            // range.getCell(long2row[url], 4).setBackground(null);
          } else {
            // there is a modlink shortlink for the modlink URL
            // there is also a file shortlink for the modlink URL
            // error: the shortlink paths do not coincide
            nbIssues++;
            // action: ask the user to correct either the file or the modlink shortlink
            logSheet.appendRow(["Error: conflict", (long2row[url]+1) + ", " + row, path + " vs. " + mappedPath, url]);
            // highlight problem shortlinks
            range.getCell(row, 4).setBackground("#ffd0d0");
            range.getCell(long2row[url], 4).setBackground("#ffd0d0");
          }
        } else if (mappedRow = long2row[url]) { // is there a local file with this URL?
          // there is a modlink shortlink for the modlink URL
          // there is a local file with this URL, without a file shortlink
          // action: use the modlink shortlink as the file shortlink
          logSheet.appendRow(["Check: File shortlink deduced", row, path, url]);
          // amend index
          range.getCell(mappedRow, 4).setValue(path).setBackground("#d0d0ff");
          // range.getCell(row, 4).setBackground("#f2f2f2");
        } else {
          // there is a modlink shortlink for the modlink URL
          // there isn't a local file with this URL
          // action: all good! (the shortlink can be created)
          // range.getCell(row, 4).setBackground("#f2f2f2");
        }
      } else if (mappedPath = long2short[url]) { // is there a file shortlink for the modlink URL?
        // there is no modlink shortlink for the modlink URL
        // but there is a file shortlink for the modlink URL
        // action: use the file shortlink as the modlink shortlink
        logSheet.appendRow(["Check: Modlink shortlink deduced", row, mappedPath, url]);
        // amend index
        range.getCell(row, 4).setValue(mappedPath).setBackground("#d0d0ff");
        // range.getCell(long2row[url], 4).setBackground(null);
      } else if (mappedRow = long2row[url]) {  // is there a local file with this URL?
        // there is no modlink shortlink for the modlink URL
        // and there is a local file with this URL, without a file shortlink
        // error: the shortlink is required, to be used in the modlink
        nbIssues++;
        // action: prompt the user for a shortlink path
        logSheet.appendRow(["Error: missing shortlink", mappedRow + ", " + row, "", url]);
        // highlight problem shortlinks
        range.getCell(row, 4).setBackground("#ffd0d0");
        range.getCell(mappedRow, 4).setBackground("#ffd0d0");
      } else {
        // there is no modlink shortlink for the modlink URL
        // and there isn't a local file with this URL
        // error: the shortlink is required, to be used in the modlink
        nbIssues++;
        // action: prompt the user for a shortlink path
        logSheet.appendRow(["Error: missing shortlink", row, "", url]);
        // highlight problem shortlink
        range.getCell(row, 4).setBackground("#ffd0d0");
      }
      */
    } else {
      // error: there is no modlink URL, the user must have deleted it!
      nbIssues++;
      // this shouldn't have happened: modlinks are created with a URL.
      logSheet.appendRow(["Error: missing URL!", row]);
    }
    
    row++;
  }
 
  /*// id-based matching
  for (var id in id2rows) {
    logSheet.appendRow(["Info: checking id", "", id]);
    lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow, 1, 1, 7).clear();
    var rows = id2rows[id];
    if (rows.length > 1) {
      var urls = {};
      for (i=0; i<rows.length; i++) {
        var row = rows[i];
        var url = range.getCell(row, 6).getValue();
        if (!urls[url]) urls[url] = true;
      }
      if (Object.keys(urls).length > 1) {
        // SpreadsheetApp.getUi().alert(Object.keys(urls).join("\n"));
        nbIssues++;
        logSheet.appendRow(["Check: Multiple URLs for the same id", rows.join(", "), "", id]);
        for (var i=0; i<rows.length; i++) {
          range.getCell(rows[i], 6).setBackground("#d0ffd0");
        }
      }
    }
  }
  */
    
  ////
  for (var url in urlMap) {
    var paths = Object.keys(urlMap[url].paths);
    var rows = urlMap[url].rows;
    if (paths.length > 1) {
      nbIssues++;
      logSheet.appendRow(["Error: shortlink path conflict", rows.join(", "), paths.join(", "), url]);
      rows.forEach(function (row) {
        range.getCell(row, 4).setBackground("#ffd0d0");
      });
    } else if (paths.length == 0) {
      if (!((rows.length == 1) && (rows[0] < modlinksRow))) {
        nbIssues++;
        logSheet.appendRow(["Error: missing shortlink", rows.join(", "), "", url]);
        rows.forEach(function (row) {
          range.getCell(row, 4).setBackground("#ffd0d0");
        });
      }
    } else {
      rows.forEach(function (row) {
        if (!range.getCell(row, 4).getValue()) {
          logSheet.appendRow(["Check: shortlink path deduced", row, paths[0], url]);
          range.getCell(row, 4).setValue(paths[0]).setBackground("#d0d0ff");
        }
      });
    }
  }  
  
  ////
  for (var path in pathMap) {
    var urls = Object.keys(pathMap[path].urls);
    var rows = pathMap[path].rows;
    if (urls.length > 1) {
      nbIssues++;
      logSheet.appendRow(["Error: multiple URLs for shortlink", rows.join(", "), urls.join("\n"), path]);
      rows.forEach(function (row) {
        range.getCell(row, 4).setBackground("#ffd0d0");
      });
    }
  }
  
  ////
  for (var id in idMap) {
    var urls = Object.keys(idMap[id].urls);
    var rows = idMap[id].rows;
    if (urls.length > 1) {
      // nbIssues++;
      logSheet.appendRow(["Check: multiple URLs for id", rows.join(", "), urls.join("\n"), id]);
      rows.forEach(function (row) {
        range.getCell(row, 6).setBackground("#d0ffd0");
      });
    //// [tmp]
    } else if (urls.length == 1) {
      rows.forEach(function (row) {
        if (!range.getCell(row, 6).getValue()) {
          logSheet.appendRow(["Check: URL deduced", row, urls[0]]);
          range.getCell(row, 6).setValue(urls[0]).setBackground("#d0d0ff");
        }
      });
    } else {
      logSheet.appendRow(["Error: no URLs for id", rows.join(", "), urls.join("\n"), id]);
    }
  }
  
  if (shortcmCheck) {
    var lastRow;
    // for every shortlink path... 
    for (var path in pathMap) {
      logSheet.appendRow(["Info: checking shortlink on short.cm", "", path]);
      lastRow = logSheet.getLastRow();
      link = expandNCCEShortlink("ncce.io/" + path);
      logSheet.getRange(lastRow, 1, 1, 7).clear();
      if (link) {
        pathMap[path].rows.forEach(function (row) {
          var url = range.getCell(row, 6).getValue();
          if (link.url == url) {
            logSheet.appendRow(["Info: shortlink exists", row, path]);
            range.getCell(row, 4).setBackground("#d0ffd0");
          } else {
            nbIssues++;
            logSheet.appendRow(["Error: shortlink exists", row, path + " (URL not the same)"]);
            range.getCell(row, 4).setBackground("#ffa0a0");
          }
        });
      }
    }
  }
  
  return nbIssues;
}  


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/// parse filename and foldername to suggest shortlink path and tags

// regular expressions that capture the resource repository naming conventions
var unitMatch = /^Y([1-9]|1[0-2]) – ([A-Za-z0-9: \-]+)(?: – ([A-Za-z0-9: \-]+)){0,1}$/;
var unitKS4Match = /^KS4 – ([A-Za-z0-9: \-]+)(?: – ([A-Za-z0-9: \-]+)){0,1}$/;
var unitComponentMatch = /^(Unit overview|Learning graph|Summative assessment|Summative assessment answers|Assessment questions|Assessment answers|Unit assessment rubric|Concept map|Learner booklet|Resource)(?:: ([A-Za-z0-9 \-]+)){0,1}(?: – ([A-Za-z0-9: \-]+)) – (?:KS4|(?:Y([1-9]|1[0-2])))$/;
var lessonMatch = /^L([1-9]|[1-9][0-9]|[1-9][0-9] (?:\&|to) [1-9][0-9]): ([A-Za-z0-9\(\)\-,!\?': ]+)$/;
var lessonComponentMatch = /^L([1-9]|[1-9][0-9]|[1-9][0-9] (?:\&|to) [1-9][0-9]) (Lesson plan|Slides|Homework|Homework solutions|Handout|Code|Resource|Rubric)(?:: ([A-Za-z0-9 \?\-\(\)]+)){0,1}(?: – ([A-Za-z0-9: \?\-]+)){0,1} – (?:KS4|(?:Y([1-9]|1[0-2])))(?:\.pdf|\.mov|\.mp4|\.wav|\.mp3|\.ogg|\.jpg|\.png|\.gif|\.bmp|\.zip|\.xlsx|\.xlsm|\.webm|\.crm){0,1}$/;
var activityComponentMatch = /^A([0-9SP]) (Worksheet|Solutions|Teacher notes|Handout|Code|Resource)(?:: ([A-Za-z0-9\?\(\)\' \-]+)){0,1}(?: – ([A-Za-z0-9\(\),!\-\?': ]+)){0,1}(?:\.pdf|\.mov|\.mp4|\.wav|\.mp3|\.ogg|\.jpg|\.png|\.gif|\.bmp|\.zip|\.xlsx|\.xlsm|\.webm|\.crm){0,1}$/;

// a mapping between the "what" field of the resource repository naming conventions and the shortlink suffix
var shortlinkMap = {
  "Unit overview": "o",
  "Learning graph": "lg",
  "Summative assessment": "saq",
  "Summative assessment answers": "saa",
  "Unit assessment rubric": "rub",
  "Assessment questions": "aq",
  "Assessment answers": "aa",
  "Concept map": "cm",
  "Learner booklet": "lm",
  "Lesson plan": "p", 
  "Slides": "s",
  "Homework": "w",
  "Homework solutions": "ws",
  "Handout": "h",
  "Code": "c",
  "Resource": "r",
  "Rubric": "rub",
  "Worksheet": "w",
  "Solutions": "s", 
  "Teacher notes": "d"
};


function suggest(filename, foldername) {
  
  var match;
  
  // try to work out a shortlink path and suggest tags
  var tags = [];
  
  // does the filename match a unit-level object?
  if (match = unitComponentMatch.exec(filename)) {
    
    // extract information from filename
    what = match[1];
    titleOpt = match[2];
    unitName = match[3];
    year = match[4];
    
    if (unitInfo.unit) {
      // determine shortlink prefix and suffix
      prefix = unitInfo.unit + "-";
      suffix = shortlinkMap[what];
      if (suffix) { if (titleOpt) suffix = suffix + titleOpt[0].toLowerCase(); } else suffix = "XXX";
    } else {
      prefix = "";
      suffix = "";
    }

    // suggest tags: what and (optional title)
    tags.push(what);
    if (titleOpt) tags.push(titleOpt);
          
    return {"path": prefix + suffix, "tags": tags};
  }
  
  // does the filename match a lesson-level object?
  if (match = lessonComponentMatch.exec(filename)) {
    
    // extract information from filename
    lesson = match[1];
    if (lesson.length > 2) lesson = lesson.substring(0,2);
    what = match[2];
    titleOpt = match[3];
    unitName = match[4];
    year = match[5];
        
    if (unitInfo.unit) {  
      // determine shortlink prefix and suffix
      var prefix = unitInfo.unit + "-" + lesson + "-";
      var suffix = shortlinkMap[what];
      if (suffix) { if (titleOpt) suffix = suffix + titleOpt[0].toLowerCase(); } else suffix = "XXX";
    } else {
      prefix = "";
      suffix = "";
    }
    
    // suggest tags: lesson number, lesson name, what and (optional title)
    tags.push("Lesson "+lesson);
    rematch = lessonMatch.exec(foldername);
    if (rematch) tags.push(rematch[2]); // lesson name, extracted from foldername
    tags.push(what);
    if (titleOpt) tags.push(titleOpt);

    return {"path": prefix + suffix, "tags": tags};
  }
    
  // does the filename match an activity-level object?
  if (match = activityComponentMatch.exec(filename)) {

    // extract information from filename
    activity = match[1];
    what = match[2];
    titleOpt = match[3];
    nameOpt = match[4];
    
    // extract additional information from foldername
    rematch = lessonMatch.exec(foldername);       
    if (rematch) {
      lesson = rematch[1]; 
      if (lesson.length > 2) lesson = lesson.substring(0,2);
    }
      else lesson = 'X';
    
    if (unitInfo.unit) {  
      // determine shortlink prefix and suffix
      var prefix = unitInfo.unit + "-" + lesson + "-a" + activity + "-";
      var suffix = shortlinkMap[what];
      if (suffix) { if (titleOpt) suffix = suffix + titleOpt[0].toLowerCase(); } else suffix = "XXX";    
    } else {
      prefix = "";
      suffix = "";
    }
    
    // suggest tags: lesson number, lesson name, what and (optional title)
    if (rematch) { // lesson number and name, extracted from foldername
      tags.push("Lesson " + rematch[1]); // lesson number
      tags.push(rematch[2]); // lesson name
    }
    tags.push(what);
    if (titleOpt) tags.push(titleOpt);
      
    return {"path": prefix + suffix, "tags": tags};
  } 
  
  return {"path": "", "tags": []};

}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/// auxillary functions

function parseRootname(foldername) {
  var match;
  if (match = unitMatch.exec(foldername))
    return {"year": match[1], "stage": getKS(match[1]), "strandOpt": match[2], "unitName": match[3]};
  
  if (match = unitKS4Match.exec(foldername))
    return {"stage": 4, "strandOpt": match[1], "unitName": match[2]};
  
  return null;
}

function getKS(year) {
  if (year < 3) return 1;
  else if (year < 7) return 2;
  else if (year < 10) return 3;
  else if (year < 13) return 4;
  else return null;
}


