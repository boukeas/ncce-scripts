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
    var targetBody = targetDoc.getBody().editAsText();
    
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

              // this is a shortlink:
              // - there's nothing to do in the target document (still a shortlink but will point to a different file)
              // - replace current shortlink **in source document** with local url (it's no longer an NCCE shortlink)
              sourceBody.setLinkUrl(link.start, link.end, currentLocalUrl);
              // [log]
              logger.appendRow(["replaced ncce.io shortlink in source", sourceDoc.getName(), "ncce.io/" + shortlink.path, currentLocalUrl]);  
              
            } else {

              // this is a normal link:
              // - there's nothing to do in the source document
              // - replace current link **in target document** with new local link          
              targetBody.setLinkUrl(link.start, link.end, newLocalUrl);
              // [log]
              logger.appendRow(["replaced link in target", targetDoc.getName(), targetBody.getText().substring(link.start, link.end+1), currentLocalUrl, newLocalUrl]);  

            }
          } else {
            if (shortlink) {
              // [log]
              if (!NCCELinks[shortlink.id]) {
                NCCELinks[shortlink.id] = currentLocalUrl;  
                // [log]
                logger.appendRow(["dry run: located ncce.io link", sourceDoc.getName(), "ncce.io/" + shortlink.path, currentLocalUrl]); 
              }
              logger.appendRow(["dry run: replaced ncce.io shortlink in source", sourceDoc.getName(), "ncce.io/" + shortlink.path, currentLocalUrl]);
            } else {
              // [log]
              logger.appendRow(["dry run: replaced link in target", targetDoc.getName(), targetBody.getText().substring(link.start, link.end+1), currentLocalUrl]);
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
  var located = false;
  var slides = SlidesApp.openById(id);
  var firstSlideElements = slides.getSlides()[0].getPageElements();
  for (var i=0; (i<firstSlideElements.length && !located); i++) {
    var element = firstSlideElements[i];
    if (element.getPageElementType() == "SHAPE") {
      var shape = element.asShape();
      var range = shape.getText().find("Save a copy");
      if (range.length > 0) {
        located = true;
        range[0].getTextStyle().setLinkUrl("https://docs.google.com/presentation/d/" + slides.getId() + "/copy");
        // [log]
        logger.appendRow(["located 'Save a copy link'", slides.getName(), slides.getId()]);
        break;
      }
    } else if (element.getPageElementType() == "TABLE") {
      var table = element.asTable();
      for (var row=0; (row < table.getNumRows() && !located); row++) {
        for (var col=0; col < table.getNumColumns(); col++) {
          try {
            var cell = table.getCell(row, col)
            var range = cell.getText().find("Save a copy");
            if (range.length > 0) {
              located = true
              range[0].getTextStyle().setLinkUrl("https://docs.google.com/presentation/d/" + slides.getId() + "/copy");
              // [log]
              logger.appendRow(["located 'Save a copy link'", slides.getName(), slides.getId()]);
              break;
            }
          } catch(error) {
            logger.appendRow(["Warning: " + error.message, slides.getName()]);
          }
        }
      }
    }
  }
}

/*
function locateLinks(editable) {
  var links = [];
  var indices = editable.getTextAttributeIndices();
  var i = 0;
  while (i < indices.length-1) {
    var offset = indices[i];    
    var url = editable.getLinkUrl(offset);
    if (url) links.push({"start": offset, "end": indices[i+1]-1, "url": url});
    i = i+1;
  }
  offset = indices[i];    
  var url = editable.getLinkUrl(offset);
  if (url) links.append({"start": offset, "end": editable.getText().length-1, "url": url});
  
  return links;
}
*/

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

var defaultUrlMatch = /(.*)\/(?:edit|view|)(?:\?usp=drivesdk){0,1}$/;

function cleanDriveUrl(url) {
  var match;
  if (match = defaultUrlMatch.exec(url)) url = match[1];
  if (match = fileUrlMatch.exec(url)) return ("https://drive.google.com/open?id=" + match[1]);
  return url;
}

/*
function updateLinks() {
  
  if (updateShortLinksFlag) {

    // use this associative array to keep track of the NCCE links already updated
    NCCELinks = {};
    
    // iterate over all document mappings's, i.e. (source, target) document pairs
    // any ncce.io short link in a target document that points to a source document, 
    // must be updated to point to the corresponding target document.  
    // note: this only pertains to "clean" links, that involve the ID of the doc, and nothing else, 
    // i.e. nothing trailing after the ID, which is quite common in Google Doc URLs. 
    
    for (var k=0; k<mappings.length; k++) {
      var mapping = mappings[k];
      if (!dryRun) {
        var linkData = updateNCCELink(mapping.sourceURL, mapping.targetURL);
        if (linkData != -1) {
          NCCELinks["https://ncce.io/" + linkData.path] = mapping.targetURL;
          // [log]
          log.appendRow(["updated clean ncce.io link", "", "ncce.io/" + linkData.path, mapping.targetURL, mapping.sourceURL]);
        }
      } else {
        var linkData = getShortLink(mapping.sourceURL);
        if (linkData != -1) {
          NCCELinks["https://ncce.io/" + linkData.path] = mapping.sourceURL;
          // [log]
          log.appendRow(["dry run: updated clean ncce.io link", "", "ncce.io/" + linkData.path, mapping.sourceURL]);
        }
      }
    }
  }
    
  // iterate over all target Docs in the 'documents' global 
  for (var i=0; i<documents.length; i++) {
    
    // open current target document
    var doc = DocumentApp.openById(documents[i]);
    
    // retrieve and iterate over all links in this document
    var links = getAllLinks(doc.getBody());
    for (var j=0; j<links.length; j++) {
      var link = links[j];
      
      // if this an NCCE link you've seen before, move on
      if (updateSLinksFlag && (NCCELinks[link.url] || NCCELinks["https://" + link.url])) continue;
      ///// CAUTION: YOU MAY NEED TO UPDATE THIS LINK TO A LOCAL CROSSLINK
      
      // iterate over all document mappings's, i.e. (source, target) document pairs
      for (var k=0; k<mappings.length; k++) {
        var mapping = mappings[k];
                  
        // check if this is an NCCE short link
        linkData = expandNCCELink(link.url);
        if (linkData != -1) { 
          var url = linkData.url;          
          // Check if the long URL links to a source document
          // (if it was a full match, we would have stumbled upon it in the first part of this function)
          if (url.indexOf(mapping.sourceId) >= 0) {
            // Match: this is an ncce.io short link in a target document that points to a source document.
            // It will either be updated to point to the corresponding target document (updateShortLinksFlag = true)
            // or it will be replaced with a local crosslink that points to the corresponding target document.

            if (!dryRun) {
              
              // this is the url of the target document that the link should point to
              var newUrl = url.replace(mapping.sourceId, mapping.targetId);
              
              if (updateShortLinksFlag) {  

                // update ncce.io short link
                var linkData = updateNCCELink(url, newUrl);
                NCCELinks["https://ncce.io/" + linkData.path] = newShortLinkUrl;                               
                // [log]
                log.appendRow(["updated ncce.io link", doc.getName(), "ncce.io/" + linkData.path, newUrl, url]);

                // CAUTION
                // replace ncce.io short link in source document with a local crosslink to another source document
                // var source = DocumentApp.openById(mapping.sourceId);
                // var newLocalUrl = link.url.replace(mapping.sourceId, mapping.targetId);
                // link.element.setLinkUrl(link.startOffset, link.endOffsetInclusive, );
                //log.appendRow(["updated cross link", doc.getName(), newUrl, link.url]);
                
              } else {

                // replace ncce.io short link                
                link.element.setLinkUrl(link.startOffset, link.endOffsetInclusive, newUrl);
                // [log]
                log.appendRow(["replaced ncce.io link with local crosslink", doc.getName(), "ncce.io/" + linkData.path, newUrl, url]);
              }
            } else {
              if (updateShortLinksFlag) {  
                NCCELinks["https://ncce.io/" + linkData.path] = url;
                // [log]
                log.appendRow(["dry run: updated ncce.io link", doc.getName(), "ncce.io/" + linkData.path, url]);
              } else {
                // [log]
                log.appendRow(["dry run: replaced ncce.io link with local crosslink", doc.getName(), "ncce.io/" + linkData.path, url]);
              }
            }
          }
        } else {
          
          // Check if the link contains the id of an original document
          if (link.url.indexOf(mapping.sourceId) >= 0) {
            // Match: this is a link in a cloned document that points to a source document.
            if (!dryRun) {
              // Update crosslink so that it points to the corresponding target document
              var newUrl = link.url.replace(mapping.sourceId, mapping.targetId);
              link.element.setLinkUrl(link.startOffset, link.endOffsetInclusive, newUrl);
              // [log]
              log.appendRow(["updated local crosslink", doc.getName(), newUrl, link.url]);
            } else {
              // [log]
              log.appendRow(["dry run: updated local crosslink", doc.getName(), link.url]);
            }
          }
        }
      }
    }
  }
}
*/


