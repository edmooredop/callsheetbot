// Replace YOUR_DROPBOX_ACCESS_TOKEN with the access token you generated in step 2
var DROPBOX_ACCESS_TOKEN = 'YOUR_DROPBOX_ACCESS_TOKEN';

function storeDropboxAppCredentials() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('DROPBOX_CLIENT_ID', 'yourclientid');
  scriptProperties.setProperty('DROPBOX_CLIENT_SECRET', 'yourclientsecret');
  scriptProperties.setProperty('DROPBOX_REFRESH_TOKEN', 'yourrefreshtoken');
}

function refreshDropboxAccessToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var clientId = scriptProperties.getProperty('DROPBOX_CLIENT_ID');
  var clientSecret = scriptProperties.getProperty('DROPBOX_CLIENT_SECRET');
  var refreshToken = scriptProperties.getProperty('DROPBOX_REFRESH_TOKEN');
  
  var tokenEndpoint = 'https://api.dropbox.com/oauth2/token';
  var requestBody = {
    'grant_type': 'refresh_token',
    'refresh_token': refreshToken,
    'client_id': clientId,
    'client_secret': clientSecret
  };
  
  var requestOptions = {
    'method': 'post',
    'payload': requestBody
  };
  
  var response = UrlFetchApp.fetch(tokenEndpoint, requestOptions);
  var jsonResponse = JSON.parse(response.getContentText());
  
  var newAccessToken = jsonResponse.access_token;
  scriptProperties.setProperty('DROPBOX_ACCESS_TOKEN', newAccessToken);
}


function processCallSheetEmails() {
  var threads = GmailApp.search('in:inbox is:unread subject:callsheet OR subject:"call sheet" has:attachment');
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var attachments = message.getAttachments();
      
      for (var k = 0; k < attachments.length; k++) {
        var attachment = attachments[k];
        var fileName = attachment.getName();
        
        if (fileName.includes("Callsheet") || fileName.includes("CS") || fileName.includes("Call sheet")) {
          refreshDropboxAccessToken();
          deleteExistingCallSheetOld();
          renameExistingCallSheet();
          saveToDropbox(attachment);
          sendPushcutNotification();
          message.markRead();

        }
      }
    }
  }
}

// Replace YOUR_PUSHCUT_API_KEY with the API Key you generated in step 3
var PUSHCUT_API_KEY = 'L4JFtSM2XlN1ahHHGEENz0ik';

function sendPushcutNotification() {
  var url = "https://api.pushcut.io/lG0iYds-RD8J9DYBe7M2l/notifications/opencallsheet";
  
  var headers = {
    "Authorization": "Bearer " + PUSHCUT_API_KEY,
    "Content-Type": "application/json"
  };
  
  var options = {
    "method": "post",
    "headers": headers
  };
  
  UrlFetchApp.fetch(url, options);
}

function renameExistingCallSheet() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var dropboxAccessToken = scriptProperties.getProperty('DROPBOX_ACCESS_TOKEN');

  var dropboxSearchApiUrl = "https://api.dropboxapi.com/2/files/search_v2";
  var searchPayload = {
    "query": "callsheet.pdf",
    "path": "/CS",
    "filename_only": true
  };
  
  var searchOptions = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + dropboxAccessToken,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(searchPayload)
  };
  
  var searchResponse = UrlFetchApp.fetch(dropboxSearchApiUrl, searchOptions);
  var searchJsonResponse = JSON.parse(searchResponse.getContentText());
  
  if (searchJsonResponse.matches.length > 0) {
    var callsheetFileId = searchJsonResponse.matches[0].metadata.id;
    var dropboxMoveApiUrl = "https://api.dropboxapi.com/2/files/move_v2";
    var movePayload = {
      "from_path": "/CS/callsheet.pdf",
      "to_path": "/CS/callsheet-old.pdf",
      "autorename": false,
      "allow_ownership_transfer": false
    };
    
    var moveOptions = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + dropboxAccessToken,
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(movePayload)
    };
    
    var moveResponse = UrlFetchApp.fetch(dropboxMoveApiUrl, moveOptions);
    var moveJsonResponse = JSON.parse(moveResponse.getContentText());
    Logger.log('Renamed existing callsheet.pdf to callsheet-old.pdf');
  } else {
    Logger.log('No existing callsheet.pdf found');
  }
}

function saveToDropbox(file) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var dropboxAccessToken = scriptProperties.getProperty('DROPBOX_ACCESS_TOKEN');
  
  var dropboxApiUrl = "https://content.dropboxapi.com/2/files/upload";
  var dropboxFilePath = "/CS/callsheet.pdf";
  
  var headers = {
    "Authorization": "Bearer " + dropboxAccessToken,
    "Content-Type": "application/octet-stream",
    "Dropbox-API-Arg": JSON.stringify({"path": dropboxFilePath, "mode": "overwrite"})
  };
  
  var options = {
    "method": "post",
    "headers": headers,
    "payload": file.getBytes()
  };
  
  var response = UrlFetchApp.fetch(dropboxApiUrl, options);
  var jsonResponse = JSON.parse(response.getContentText());
  
  if (jsonResponse.name && jsonResponse.path_display) {
    Logger.log('File saved to Dropbox: %s', jsonResponse.path_display);
  } else {
    Logger.log('Error saving file to Dropbox:');
    Logger.log(jsonResponse);
  }
}


function deleteExistingCallSheetOld() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var dropboxAccessToken = scriptProperties.getProperty('DROPBOX_ACCESS_TOKEN');

  var dropboxSearchApiUrl = "https://api.dropboxapi.com/2/files/search_v2";
  var searchPayload = {
    "query": "callsheet-old.pdf",
    "path": "/CS",
    "filename_only": true
  };
  
  var searchOptions = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + dropboxAccessToken,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(searchPayload)
  };
  
  var searchResponse = UrlFetchApp.fetch(dropboxSearchApiUrl, searchOptions);
  var searchJsonResponse = JSON.parse(searchResponse.getContentText());
  
  if (searchJsonResponse.matches.length > 0) {
    var callsheetOldFileId = searchJsonResponse.matches[0].metadata.id;
    var dropboxDeleteApiUrl = "https://api.dropboxapi.com/2/files/delete_v2";
    var deletePayload = {
      "path": "/CS/callsheet-old.pdf"
    };
    
    var deleteOptions = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + dropboxAccessToken,
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(deletePayload)
    };
    
    var deleteResponse = UrlFetchApp.fetch(dropboxDeleteApiUrl, deleteOptions);
    var deleteJsonResponse = JSON.parse(deleteResponse.getContentText());
    Logger.log('Deleted existing callsheet-old.pdf');
  } else {
    Logger.log('No existing callsheet-old.pdf found');
  }
}


