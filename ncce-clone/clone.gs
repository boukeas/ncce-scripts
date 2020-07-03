// Clones a folder (source) and all its contents into a new folder (target).
// Comments in documents are not preserved in the copies.

// Features:
// - A spreadsheet log file is created in your root MyDrive folder. You can monitor the script process as it runs by viewing the log.
// - In the target, all links in all Documents that point to source files can be updated to point to target files (updateXLinksFlag).
// - All ncce.io links that point to source files can be updated to point to target files (updateSLinksFlag).
// - You can replace NCCE links with local crosslinks. This would allow working on local copies.
// - You can simulate the cloning process without actually performing it (dryRun).
// - You can exclude files and folders from cloning by specifying their id's or specific prefixes (excludedIds and excludedPrefixes). 
// - You can set specific permissions, different from the original ones, for the cloned files and folders. (defaultAccess and defaultPermission).

// Todo's:
// - Deactivate Start while cloning in progress
// - Add “save a copy” link diagnostics to Dry run and tally
// - Add a final line to the log with Time, Date, statistics
// - refine cleanDriveUrl
// - Add additional optional parameters to sidebar
// - keep current rights if alternatives are not specified  

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Set the variables below to suit your needs. 
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// This is the source folder to be cloned. You need to specify its id. 
// var sourceId = '1yifHge1rlcUhUcF-b3QrToNxqnOWYki-';
// var sourceId = "1IVSwANIk-wtafgqBkUGp7lS7nRNxsUKH";
var sourceId;

// This is the parent of the target folder. The target folder will be created inside the parent.
// If you don't specify it's id, the target folder will be created under your home Drive.
// You can set the parentId to the sourceId, to create the clone under the source folder.
// var parentId = sourceId;
var parentId;
// var parentId = '1Zq9Be33yC6xwoIaJLOpsRLLPu--gTwR9';

// This is the prefix added to the target folder (and you can always rename the clone later)
// var prefix = "Clone of ";
var prefix;

// The ids of files or folders included in this array will *not* be cloned.
var excludedIds = [];
// Folders whose names start with any of the prefixes included in this list will *not* be cloned.
var excludedPrefixes = ["old", "[Ignore]"];

// Set this flag depending on where you want the ncce.io shortlinks to redirect to.
//
// updateShortlinksFlag               | true                        | false                       |
// ----------------------------------------------------------------------------------------------
// ncce.io shortlinks in the source | unchanged                   | changed to local crosslinks |
// ncce.io shortlinks in the target | changed to local crosslinks | unchanged                   |

// var keepShortlinksFlag = true;
var updateShortlinksFlag;

// Set this flag to True if you want to *test* the cloning process without actually creating any clones.
// Review the log after the script completes execution.
// dryRun = false;
var dryRun;

// global variable to handle short.cm authorisation
// get your key from https://app.short.cm/users/integrations/api-key 
authorisationKey = "2qVUc74P6LxkByOz";

// Default access and permission rights for cloned documents.
var defaultAccess;
// var defaultAccess = DriveApp.Access.DOMAIN;
var defaultPermission;
// var defaultPermission = DriveApp.Permission.EDIT;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// globals
var mapping = {};
var documents = [];
var logger = SpreadsheetApp.getActiveSheet();

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function clone_init() {
    
  if (!sourceId) {
    Logger.log("Error: You need to set the sourceId variable.");
    return;
  }
  
  logger.clearContents();

  // look for source folder
  var source = DriveApp.getFolderById(sourceId);
  var sourceName = source.getName();
  
  // set name of target folder: prefix + source + timestamp
  // var timestamp = Date.now();
  // var targetName = prefix + sourceName + " " + timestamp;
  var targetName = prefix + sourceName;
  var target;
  excludedPrefixes.push(prefix);

  // Create log spreadsheet in MyDrive
  // var logBook = SpreadsheetApp.create("Clone log " + sourceName + " " + timestamp);
  // log = logBook.getActiveSheet();
  
  if (!dryRun) {
    
    if (parentId) {
      // look for parent of target folder: this will throw an exception if the id does not exist
      parent = DriveApp.getFolderById(parentId);
      // create target (inside parent)
      target = parent.createFolder(targetName);
      // [log]
      logger.appendRow(["created target folder in " + parent.getName(), target.getName(), target.getId(), source.getId()]);
    } else {
      // create target (inside home drive folder)
      target = DriveApp.createFolder(targetName);
      // [log]
      logger.appendRow(["created target folder in MyDrive", target.getName(), target.getId(), source.getId()]);
    }
      
    // new target folder is added to the list of folders that should not be cloned
    // because it may be a descendant of the source folder (avoid infinite recursion) 
    excludedIds.push(target.getId());
    
    // [Alt] Cannot use .setSharing in shared drives
    try {
      target.setSharing(defaultAccess, defaultPermission); 
    } catch(error) {
      logger.appendRow(["error: " + error.message, target.getName()]);
    }
    
  } else {

    if (parentId) {
      // look for parent of target folder: this will throw an exception if the id does not exist
      parent = DriveApp.getFolderById(parentId);
      // [log]
      logger.appendRow(["dry run: created target folder in " + parent.getName(), source.getName(), source.getId()]);      
    } else {
      // [log]
      logger.appendRow(["dry run: created target folder in MyDrive", source.getName(), source.getId()]);      
    }

  }
    
  // clone
  clone(source, target);
  
  // update links
  if (updateShortlinksFlag) updateShortlinks(); else expandShortlinks();
}

function clone(source, target) {
    
  // for all files in source  
  var files = source.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    
    // check for file exclusions
    if (excludedIds.indexOf(file.getId()) >= 0) {
      // [log]
      logger.appendRow(["excluded file (by id)", file.getName(), file.getId()]);
      continue;
    }
    if (prefixMembership(file.getName(), excludedPrefixes)) {
      // [log]
      logger.appendRow(["excluded file (by prefix)", file.getName(), file.getId()]);
      continue;
    }
    
    if (!dryRun) {
      
      try {
        // make a copy of the file in the target folder
        copy = file.makeCopy(file.getName(), target);
        // record a mapping from the id of the source file to the id of its copy (so we know what to update in the links)
        mapping[file.getId()] = copy.getId();
        // [log]
        logger.appendRow(["copied file", copy.getName(), copy.getId(), file.getId()]);
      } catch(error) {
        logger.appendRow(["Unable to create a copy: " + error.message, file.getName(), target.getName()]);
      }
      
      try {
        // apply default permission settings
        copy.setSharing(defaultAccess, defaultPermission);
      } catch(error) {
        logger.appendRow(["error: " + error.message, copy.getName()]);
      }
        
      // if the file is a document:
      // record the id of the copy (so we know which files will need updated links)
      // [caution] Links will *only* be updated in Docs. Save a copy links will be created for Docs and Sheets.
      if (file.getMimeType() == "application/vnd.google-apps.document") {
        documents.push(file.getId());
        updateSaveACopyDoc(copy.getId());
      } else if (file.getMimeType() == "application/vnd.google-apps.presentation") {
        updateSaveACopyPresentation(copy.getId());
      }
      
    } else {
      mapping[file.getId()] = file.getId();
      if (file.getMimeType() == "application/vnd.google-apps.document") documents.push(file.getId());
      // [log]
      logger.appendRow(["dry run: copied file", file.getName(), file.getId()]);
    }
  }
    
  // for all folders in source
  var subFolders = source.getFolders();
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    
    // check for folder exclusions
    if (excludedIds.indexOf(subFolder.getId()) >= 0) {
      // [log]
      logger.appendRow(["excluded folder (by id)", subFolder.getName(), subFolder.getId()]);
      continue;
    }
    if (prefixMembership(subFolder.getName(), excludedPrefixes)) {
      // [log]
      logger.appendRow(["excluded folder (by name)", subFolder.getName(), subFolder.getId()]);
      continue;
    }

    // create and copy folder contents (recursive call)
    var folderName = subFolder.getName();
    if (!dryRun) {
      var targetFolder = target.createFolder(folderName);
      // [Alt] Cannot use .setSharing in shared drives
      try {
        targetFolder.setSharing(defaultAccess, defaultPermission);
      } catch(error) {
        logger.appendRow(["error: " + error.message, targetFolder.getName()]);
      }
      // [log]
      logger.appendRow(["created folder", targetFolder.getName(), targetFolder.getId(), subFolder.getId()]);
    } else {
      // [log]
      logger.appendRow(["dry run: created folder", subFolder.getName(), subFolder.getId()]);
    }
    
    // recursive call
    clone(subFolder, targetFolder);
  }    
}v
