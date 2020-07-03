var authorisationKey = "2qVUc74P6LxkByOz";

function getNCCEShortlink(URL) {

  // request and return data object for long URL (id and path)
  headers = { 'authorization': authorisationKey };
  parameters = { 'headers': headers };
  try {
    response = UrlFetchApp.fetch('https://api.short.cm/links/by-original-url?domain=ncce.io&originalURL=' + encodeURIComponent(URL), parameters);
    if (response.getResponseCode() == 200) {
      content = JSON.parse(response.getContentText()); 
      return {
        "id": content.id.toString(),
        "path": content.path
      };
    } else return null;
  } catch(error) {
    Logger.log(URL + " " + error);
    return null;
  }
}


function expandNCCEShortlink(URL) {

  // check if URL is an NCCE shortlink and extract path  
  var start = URL.indexOf("ncce.io/");
  if (start < 0) return null;
  var path = URL.substring(start+8);
  
  // request and return data object for path (id and long URL)
  headers = { 'authorization': authorisationKey };
  parameters = { 'headers': headers };
  try {
    response = UrlFetchApp.fetch('https://api.short.cm/links/expand?domain=ncce.io&path=' + path, parameters);
    if (response.getResponseCode() == 200) {
      content = JSON.parse(response.getContentText()); 
      return {
        "id": content.id.toString(),
        "path": content.path,
        "url": content.originalURL
      };
    } else return null;
  } catch(error) {
    Logger.log(URL + " " + error);
    return null;
  }
}
 
function updateNCCEShortlink(id, newURL) {
  
  // update short link (using id)
  var payload = { 'originalURL': newURL };
  headers = { 
    'authorization': authorisationKey,
    'content-type': "application/json"
  };
  parameters = {
    'method': "post",  
    'headers': headers,
    'payload': JSON.stringify(payload)
  };
  try {
    response = UrlFetchApp.fetch('https://api.short.cm/links/' + id, parameters);
    if (response.getResponseCode() == 200) {
      content = JSON.parse(response.getContentText());
      return {
        "id": id,
        "path": content.path.toString()
      };
    } else return null;
  } catch(error) {
    Logger.log(error);
    return null;
  }
}

function bulkNCCEShortlinks(links) {
  var payload = { 
    "domain": "ncce.io",
    "links": links
  };
  headers = { 
    'authorization': authorisationKey,
    'content-type': "application/json"
  }; 
  parameters = {
    'method': "post",  
    'headers': headers,
    'payload': JSON.stringify(payload)
  };
  try {
    var response = UrlFetchApp.fetch('https://api.short.cm/links/bulk', parameters);
    if (response.getResponseCode() == 200) {
      return JSON.parse(response.getContentText());
    } else return null;
  } catch(error) {
    SpreadsheetApp.getUi().alert(error);
    return null;
  }
}

/*
function updateNCCEShortlink(currentURL, newURL) {
  
  // request id for currentURL short link
  headers = { 'authorization': authorisationKey };
  parameters = { 'headers': headers };
  try {
    response = UrlFetchApp.fetch('https://api.short.cm/links/by-original-url?domain=ncce.io&originalURL=' + encodeURIComponent(currentURL), parameters);
    if (response.getResponseCode() == 200) {
      content = JSON.parse(response.getContentText());
      id = content.id.toString(); 
    } else 
      return null;
  } catch(error) {
    Logger.log(currentURL + " " + error);
    return null;
  }
  
  // update short link (using id)
  var payload = { 'originalURL': newURL };
  headers = { 
    'authorization': "2qVUc74P6LxkByOz",
    'content-type': "application/json"
  };
  parameters = {
    'method': "post",  
    'headers': headers,
    'payload': JSON.stringify(payload)
  };
  try {
    response = UrlFetchApp.fetch('https://api.short.cm/links/' + id, parameters);
    if (response.getResponseCode() == 200) 
      return {
        "id": id,
        "path": content.path.toString()
      };
    else return null;
  } catch(error) {
    Logger.log(error);
    return null;
  }
}
*/
