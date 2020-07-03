///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/// create shortlinks

function create_init(unitTags) {
       
  var indexSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index");
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Create").activate();
  
  // clear log and create header row
  logSheet.clear();
  logSheet.appendRow(["Message", "Path", "Title", "URL"]);
  logSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  var path;
  var pathMap = {};
  
  var nbRows = indexSheet.getLastRow();
  range = indexSheet.getRange(1, 1, nbRows, 7);
  for (var row=2; row<=nbRows; row++) {
    if ((path = range.getCell(row,4).getValue()) && !pathMap[path]) {
      if (!(title = range.getCell(row,3).getValue())) title = range.getCell(row,2).getValue();
      pathMap[path] = {
        "title": title,
        "tagStr": range.getCell(row,5).getValue(),
        "url": range.getCell(row,6).getValue()
      }
      // logSheet.appendRow(["DEBUG read", path, pathMap[path].title, pathMap[path].tagStr, pathMap[path].url]); 
    }
  }
  
  var links = [];
  for (path in pathMap) {
    var tags = unitTags.concat(
                 pathMap[path].tagStr.split("|").
                 map(Function.prototype.call, String.prototype.trim));
    
    links.push({
      "originalURL": pathMap[path].url,
      "path": path,
      "title": pathMap[path].title,
      "tags": tags});
    
    var linkRecord = links[links.length-1];
    // logSheet.appendRow(["DEBUG create", linkRecord.path, linkRecord.title, linkRecord.tags.join(" | "), linkRecord.originalURL]); 
  }
  
  if (links.length) {
    var responses = bulkNCCEShortlinks(links);
    for (var i = 0; i<responses.length; i++) {
      var response = responses[i];
      if (response.success) {
        logSheet.appendRow(["Successful", response.title, response.path, response.originalURL]);
      } else {
        logSheet.appendRow(["Error: " + response.status, links[i].title, links[i].path, links[i].originalURL]);
      }
    }
  } else 
    return null;
}

function replace_init(dryRun) {
  
  var indexSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Index");
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Replace").activate();
  
  // clear log and create header row
  logSheet.clear();
  logSheet.appendRow(["File", "Action"]);
  logSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  
  var path, url;
  var urlMap = {};
  
  var nbRows = indexSheet.getLastRow();
  range = indexSheet.getRange(1, 1, nbRows, 7);
  for (var row=2; row<=nbRows; row++) {
    if ((path = range.getCell(row,4).getValue()) && (url = range.getCell(row,6).getValue()) && !urlMap[url]) {
      urlMap[url] = path; // [CAUTION] this *assumes* no duplicates
    }
  }
  
  // iterate over documents and call replaceModlinks()
  
  range = indexSheet.getRange(1, 3, nbRows, 5);
  // go through all file rows  
  row = 2;
  while (row <= nbRows && !range.getCell(row,1).getValue()) {
    var id = range.getCell(row,5).getValue();
    var file = DriveApp.getFileById(id);
    if (file.getMimeType() == "application/vnd.google-apps.document") {
      var results = replaceModlinks(DocumentApp.openById(id).getBody().editAsText(), urlMap, dryRun);
      logSheet.appendRow([file.getName(), "Located " + results.located + ", replaced " + results.replaced + " with shortlinks, " + results.inconsistent + " were inconsistent"]);
      // if (results.replaced) logSheet.appendRow([file.getName(), "Replaced " + results.replaced + " links with shortlinks"]);
      // if (results.inconsistent) logSheet.appendRow([file.getName(), results.inconsistent + " existing shortlinks were inconsistent"]);
    }  
    row++;
  }
}

function replaceModlinks(editable, mappings, dryRun) {
  var located = 0;
  var replaced = 0;
  var inconsistent = 0;
  // var contested = [];
  var offset = 0;
  var path;
  locateModlinks(editable).forEach(function (modlink) {
    located++;
    if (path = mappings[cleanDriveUrl(modlink.url)]) {  // is the target url in the link mapped to a shortlink path?
      
      replaced++;
      if (!dryRun) {
        // replace link
        editable.setLinkUrl(modlink.start + offset, modlink.end + offset, "https://ncce.io/" + path);
      }
      
      if (modlink.path) { // has a shortlink path also been provided in brackets?
        if (path != modlink.path) {
          inconsistent++; // check: are these two shortlink paths identical?
          if (!dryRun) {
            // replace path in brackets
            editable.deleteText(modlink.bstart + offset + 1, modlink.bstart + offset + 8 + modlink.path.length);
            editable.insertText(modlink.bstart + offset + 1, "ncce.io/" + path);
            offset = offset + path.length - modlink.path.length;
          }
        }
      } else {
        if (!dryRun) {
          // insert new path in brackets
          editable.insertText(modlink.bstart + offset + 1, "ncce.io/" + path);
          offset = offset + path.length + 8;
        }      
      }
    }
  });
  
  return {"located": located, "replaced": replaced, "inconsistent": inconsistent};
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function expandShortlinks() {
  // iterate over all target Docs in the 'documents' global 
  documents.forEach(function(id) {
    
    // open current target document
    var targetDoc = DocumentApp.openById(mapping[id]);
    var targetBody = targetDoc.getBody().editAsText();
    
    // retrieve and iterate over all links in this document
    locateLinks(targetBody).forEach(function(link) {
      var currentLocalUrl;
            
      // check if this is an NCCE short link
      shortlink = expandNCCEShortlink(link.url);
      if (shortlink) {  
        currentLocalUrl = shortlink.url;
      } else {
        currentLocalUrl = link.url;
      }
      
      // Check if the local URL is a drive url (and extract document id)
      var sourceId = isDriveUrl(currentLocalUrl);
      if (sourceId) {
        // Check if the source id has been been mapped to a target id
        var targetId = mapping[sourceId];
        if (targetId) {
          if (!dryRun) {
            
            // this is the url of the target document that the link should point to
            var newLocalUrl = currentLocalUrl.replace(sourceId, targetId);
            
            // replace current link **in target document**      
            // (it doesn't matter if it used to be an NCCE link or not, now it's definitely not)
            targetBody.setLinkUrl(link.start, link.end, newLocalUrl);
            // [log]
            logger.appendRow(["replaced link in target", targetDoc.getName(), link.url == currentLocalUrl ? "" : link.url, currentLocalUrl, newLocalUrl]);  
          
          } else {
            // [log]
            logger.appendRow(["dry run: replaced link in target", targetDoc.getName(), link.url == currentLocalUrl ? "" : link.url, currentLocalUrl]);
          }
        }
      }
    });
  });
}

function updateShortlinks() {
  // use this associative array to keep track of the NCCE links already updated
  NCCELinks = {};
  
  // iterate over all document mappings's, i.e. (source, target) document pairs:
  // any ncce.io short link in a target document that points to a source document, 
  // must be updated to point to the corresponding target document.  
  // note: this only pertains to "clean" links, that involve the ID of the doc, and nothing else, 
  // i.e. nothing trailing after the ID, which is quite common in Google Doc URLs. 
  
  for (var sourceId in mapping) {
    // use the source file id to retrieve (and clean) its URL.
    var sourceUrl = cleanDriveUrl(DriveApp.getFileById(sourceId).getUrl());
    // retrieve shortlink (if there is actually one pointing to the sourceUrl)
    var shortlink = getNCCEShortlink(sourceUrl);
    if (shortlink) {
      if (!dryRun) {
        // use the target file id to retrieve (and clean) its URL.
        var targetUrl = cleanDriveUrl(DriveApp.getFileById(mapping[sourceId]).getUrl());
        NCCELinks[shortlink.id] = targetUrl;
        // [log]
        logger.appendRow(["located clean ncce.io link", "", "ncce.io/" + shortlink.path, sourceUrl, targetUrl]);
      } else {
        NCCELinks[shortlink.id] = sourceUrl;
        // [log]
        logger.appendRow(["dry run: located clean ncce.io link", "", "ncce.io/" + shortlink.path, sourceUrl]);
      }
    } 
  }
  
  // iterate over all target Docs in the 'documents' global 
  documents.forEach(function(id) {
    
    // open current target document
    var sourceDoc = DocumentApp.openById(id);
    var sourceBody = sourceDoc.getBody().editAsText();
    var targetDoc = DocumentApp.openById(mapping[id]);
    var targetBody = sourceDoc.getBody().editAsText();
    
    // retrieve and iterate over all links in this document
    locateLinks(targetBody).forEach(function(link) {
      var currentLocalUrl;

      // check if this is an NCCE short link
      var shortlink = expandNCCEShortlink(link.url);
      if (shortlink) { 
        currentLocalUrl = shortlink.url;
      } else {
        currentLocalUrl = link.url;
      }
      
      // Check if the local URL is a drive url (and extract document id)
      var sourceId = isDriveUrl(currentLocalUrl);
      if (sourceId) {
        // Check if the id has been been mapped to a target id
        var targetId = mapping[sourceId];
        if (targetId) {
          if (!dryRun) {
            
            // this is the url of the target document that the link should point to
            var newLocalUrl = currentLocalUrl.replace(sourceId, targetId);
            
            if (shortlink) {    

              if (!NCCELinks[shortlink.id]) {
                NCCELinks[shortlink.id] = newLocalUrl;                               
                // [log]
                logger.appendRow(["located ncce.io link", sourceDoc.getName(), "ncce.io/" + shortlink.path, currentLocalUrl, newLocalUrl]); 
              }

              // replace current link **in source document**     
              // (it's no longer an NCCE shortlink)
              sourceBody.setLinkUrl(link.start, link.end, currentLocalUrl);
              // [log]
              logger.appendRow(["replaced link in source", sourceDoc.getName(), "ncce.io/" + shortlink.path, currentLocalUrl]);  
              
            } else {

              // replace current link **in target document**          
              targetBody.setLinkUrl(link.start, link.end, newLocalUrl);
              // [log]
              logger.appendRow(["replaced link in target", targetDoc.getName(), "", currentLocalUrl, newLocalUrl]);  

            }
          } else {
            if (shortlink) {
              // [log]
              if (!NCCELinks[shortlink.id]) {
                NCCELinks[shortlink.id] = currentLocalUrl;  
                // [log]
                logger.appendRow(["dry run: located ncce.io link", sourceDoc.getName(), "ncce.io/" + shortlink.path, currentLocalUrl]); 
              }
              logger.appendRow(["dry run: replaced link in source", sourceDoc.getName(), "ncce.io/" + shortlink.path, currentLocalUrl]);
            } else {
              // [log]
              logger.appendRow(["dry run: replaced link in target", targetDoc.getName(), "", currentLocalUrl]);
            }
          }
        }
      }
    });
  });
  
  if (!dryRun) {
    for (var id in NCCELinks) {
      var shortlink = updateNCCEShortlink(id, NCCELinks[id]);
      // [log]
      logger.appendRow(["updated ncce.io link", " ", "ncce.io/" + shortlink.path, " ", NCCELinks[id]]); 
    }
  }
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/// auxillary functions


var bracketMatch = /\((?:ncce.io\/([\A-Za-z0-9\-]+)){0,1}\)/g;

function locateModlinks(editable) {
  // locates and returns all links followed by a pair of brackets, 
  // possibly containing an NCCE shortlink
  var modlinks = [];
  
  // regular expression to match all brackets containing an NCCE shortlink
  var match = bracketMatch.exec(editable.getText());
  // get list of links
  var links = locateLinks(editable);
  var linkIndex = 0;

  // go through the links and the bracket matches simultaneously
  // advancing either one or the other
  // we are interested in the cases where the link is followed by the brackets
  while (match && linkIndex < links.length) {
    var link = links[linkIndex];
    var linkEnd = link.end;
    var shortlinkStart = match.index;
    if (shortlinkStart < linkEnd + 2) {
      match = bracketMatch.exec(editable.getText());
    } else if (shortlinkStart > linkEnd + 2) {
      linkIndex++;
    } else {
      // match!
      //Logger.log(editable.getText().substring(link.start, link.end+1) + " " + match[1]);
      modlinks.push({
        "start": link.start, 
        "end": link.end, 
        "text": editable.getText().substring(link.start, link.end+1),
        "url": link.url,
        "bstart": match.index,
        "path": match[1]
      });
      match = bracketMatch.exec(editable.getText());
      linkIndex++;
    }
  }

  return modlinks;
}

function locateLinks(editable) {
  var links = [];
  var indices = editable.getTextAttributeIndices();
  var i = 0;
  while (i < indices.length-1) {
    var offset = indices[i];    
    var url = editable.getLinkUrl(offset);
    if (url) {
      if (links.length && (lastLink = links[links.length-1]).url == url && lastLink.end == offset-1) {
        // merge with previous link, they are consecutive and point to the same url
        lastLink.end = indices[i+1]-1;
      } else {
        links.push({"start": offset, "end": indices[i+1]-1, "url": url});
      }
    }
    i = i+1;
  }
  offset = indices[i];    
  var url = editable.getLinkUrl(offset);
  if (url) links.push({"start": offset, "end": editable.getText().length-1, "url": url});
  
  links.forEach(function (link) {
    Logger.log("from " + (link.start) + " to " + (link.end) + " " + editable.getText().substring(link.start, link.end+1) + " > " + link.url);
  })
  
  return links;
}


var driveUrlMatch = /^(?:https:\/\/){0,1}docs\.google\.com\/(?:document|presentation|spreadsheets|drawings)\/d\/([\w-]+)(?:\/[\w-\?\.=#\/]+){0,1}/;
var shareUrlMatch = /^(?:https:\/\/){0,1}drive\.google\.com\/open\?id=([\w-]+)$/
var fileUrlMatch = /^(?:https:\/\/){0,1}drive\.google\.com\/[\w-\.\/]+\/file\/d\/([\w-]+)$/;

function isDriveUrl(url) {
  var match;
  if (match = driveUrlMatch.exec(url)) return match[1]; else 
  if (match = shareUrlMatch.exec(url)) return match[1]; else 
  if (match = fileUrlMatch.exec(url)) return match[1]; else 
  return null;
}

function prefixMembership(str, prefixes) {
  for (i = 0; i<prefixes.length; i++) {
    if (str.indexOf(prefixes[i]) == 0) {
      return true;
    }
  }
  return false;
}

/*
function cleanDriveUrl(url) {
  // remove "/edit?usp=drivesdk" from the end of the string
  return url.substring(0,url.length-18);
}
*/

var defaultUrlMatch = /(.*)\/(?:edit|view|)(?:\?usp=drivesdk|usp=sharing){0,1}$/;

function cleanDriveUrl(url) {
  var match;
  if (match = defaultUrlMatch.exec(url)) url = match[1];
  if (match = fileUrlMatch.exec(url)) return ("https://drive.google.com/open?id=" + match[1]);
  return url;
}

function compare(url1, url2) {
  if (defaultUrlMatch.exec(url1) && fileUrlMatch.exec(url2)) return url2;
  if (defaultUrlMatch.exec(url2) && fileUrlMatch.exec(url1)) return url1;
  return false;
}

function prefixMembership(str, prefixes) {
  for (i = 0; i<prefixes.length; i++) {
    if (str.indexOf(prefixes[i]) == 0) {
      return true;
    }
  }
  return false;
}

