function onInstall() {
  onOpen();
}

function onOpen() {
  var menuEntries = [
    {name: "Generate Bibliography", functionName: "generateBibliography"},
    {name: "Configure", functionName: "configure"}
  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Exciting", menuEntries);
}

var bibliography;
var updated;
var results = {};
var known = {};

function generateBibliography(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var docName = sheet.getName().replace(/ - References$/, "");

  var file = findDocumentInFolders(docName, DocsList.getFileById(sheet.getId()).getParents());
  if (!file) return false;

  var tmpFile = DocsList.copy(file, docName + " - Formatted Citations");

  var tmpDoc = DocumentApp.openById(tmpFile.getId());
  var body = tmpDoc.getActiveSection();
  bibliography = body.appendTable();
  bibliography.setBorderColor("#FFFFFF");

  known = loadItems();
  updated = Math.round(Date.now() / 1000); // current UNIX timestamp

  try {
    findResults(body);
  }
  catch (e){
    return false;
  }

  var counter = 1;
  for (var resultsKey in results){
    var item = results[resultsKey];
    Logger.log(item.text);


    var replacement = "[" + counter + "]";

    body.replaceText(item.text, replacement);

    /*
    // TODO: could match false positives, but haven't got the offsets of the replacements
    var positions = getPositions(body, replacement);
    for (var i = 0; i < positions.length; i++){
      var position = positions[i];
      // TODO: bookmark links to the bibliography items
      // http://code.google.com/p/google-apps-script-issues/issues/detail?id=803
      position.element.setLinkUrl(position.startOffset, position.startOffset + replacement.length, createUrl(item.key));
    }
    */

    addBibItem(counter, item.data);
    counter++;
  }

  tmpDoc.saveAndClose();
  var pdf = tmpDoc.getAs("application/pdf").getBytes();
  tmpFile.setTrashed(true);

  saveItems(results);

  MailApp.sendEmail(
    Session.getUser().getEmail(),
    "Excited document!",
    "Your document with formatted citations is attached",
    {attachments: [{fileName: "exciting.pdf", mimeType: "application/pdf", content: pdf}]}
    );

  Browser.msgBox("Done!");
}

function findDocumentInFolders(docName, folders){
  if (!folders.length) {
    folders = DocsList.getFolders();
  }

  var file;
  for (var i = 0; i < folders.length; i++){
    var files = folders[i].getFilesByType("document");
    for (var j = 0; j < files.length; j++){
      file = files[j];
      if (file.getName() == docName) {
        return file;
      }
    }
  }
}

function findResults(node){
  var result = null;
  
  do{
    result = node.findText("{{cite:.+?}}", result);
    if (!result) break;
      
    var startOffset = result.getStartOffset();
    var endOffset = result.getEndOffsetInclusive();
    
    var text = result.getElement().getText().substr(startOffset, (endOffset - startOffset) + 1);

    var match = text.match(/^\{\{cite:(\w+:.+?)\}\}$/i);
    if (!match) return false;
    
    var key = match[1];

    //var resultKey = Utilities.base64Encode(key);
    var resultKey = Base64.encode(key);
    if (typeof results[resultKey] != "undefined") continue;

    results[resultKey] = {
      text: text,
      key: key,
      data: fetchData(key)
    };
    
  } while (result);
}

/*
function getPositions(body, text){
  var result = null;
  var positions = [];
  
  do {
    result = body.findText(text, result);
    if (!result) break;

    positions.push({ element: result.getElement(), startOffset: result.getStartOffset() });
  } while (result);
  
  return positions;
}
*/

function createUrl(key){
  var match = key.match(/^(\w+):(.+)/);
  if (!match) return false;

  var type = match[1];
  var id = match[2];

  switch (type){
  case "doi":
    return "http://dx.doi.org/" + id;

  case "http":
  case "https":
  case "ftp":
    return key;

  default:
    return false;
  }
}

function fetchData(id){
  var data = fetchLocal(id);
  if (data) return data;

  var type, match = id.match(/^(\w+):(.+)/);
  if (match){
    type = match[1].toLowerCase();
    id = match[2].replace(/\|.+/, "");
  }

  Logger.log(id);
  if (!id) return false;

  //type = "test";

  switch (type){
  case "doi":
    return fetchByDoi(id);

  case "mendeley":
    return fetchMendeleyById(id);

  case "test":
    return { "title": "Test Citation" };

  default:
    return fetchByUrl(type + ":" + id);
  }
}

function addBibItem(citation, data){
  Logger.log(data.toSource());
  var row = bibliography.appendTableRow();
  row.appendTableCell().appendParagraph(citation);

  var cell = row.appendTableCell();
  Logger.log(cell.getAttributes());
  var title = cell.appendParagraph(data["title"] + ". " + data["container-title"] + "  (" + data["date"] + ")"); // TODO: CiteProc
}

function fetchLocal(key){
  var resultKey = Base64.encode(key);
  if (typeof known[resultKey] != "undefined"){
    Logger.log("found " + key);
    return known[resultKey];
  }
}

function fetchByDoi(doi){
  var data;
  //if (isConfigured()) data = fetchMendeleyByDoi(doi);
  if (!data) data = fetchCrossRefByDoi(doi);
  return data;
}

function fetchCrossRefByDoi(doi){
  Logger.log("Fetching from CrossRef by DOI: " + doi);

  var result = UrlFetchApp.fetch("http://dx.doi.org/" + doi, { "headers": {"Accept": "application/json" } });
  if (result.getResponseCode() != 200) return false;

  var data = Utilities.jsonParse(result.getContentText());
  if (!data.feed || !data.feed.entry) return false;

  return normaliseCrossRef(data.feed.entry);
}

function fetchMendeleyByDoi(doi){
  setupOAuth();

  Logger.log("Fetching from Mendeley by DOI: " + doi);

  var options = { "method": "GET", "oAuthServiceName": "mendeley", "oAuthUseToken": "never" };
  var result = UrlFetchApp.fetch("http://api.mendeley.com/oapi/documents/details/" + encodeURIComponent(doi.replace("/", encodeURIComponent("/"))) + "/?type=doi", options);
  if (result.getResponseCode() != 200) return false;

  var data = Utilities.jsonParse(result.getContentText());
  if (!data) return false;

  return normaliseMendeley(data);
}

function fetchMendeleyById(id){
  setupOAuth();

  Logger.log("Fetching from Mendeley by id: " + id);

  var options = { "method": "GET", "oAuthServiceName": "mendeley", "oAuthUseToken": "always" };
  var result = UrlFetchApp.fetch("http://api.mendeley.com/oapi/library/" + encodeURIComponent(id) + "/", options);
  if (result.getResponseCode() != 200) return false;

  var data = Utilities.jsonParse(result.getContentText());
  Logger.log(data);
  if (!data) return false;

  return normaliseMendeley(data);
}

function fetchByUrl(url){
  Logger.log("Fetching URL: " + url);

  var result = UrlFetchApp.fetch(url, { "headers": {"Accept": "application/json" } });
  if (result.getResponseCode() != 200) return false;

  var data = Utilities.jsonParse(result.getContentText());
  if (!data) return false;

  return normaliseGeneric(data.feed.entry);
}

function normaliseCrossRef(data){
  Logger.log(data.toSource());

  article = data["pam:message"]["pam:article"];

  var item = {
    "title": article["dc:title"],
    "doi": article["prism:doi"],
    "date": article["prism:publicationDate"],
    "authors": article["dc:creator"],
    "start-page": article["prism:startingPage"],
    "end-page": article["prism:endingPage"],
    "container-issn": article["prism:issn"],
    "container-title": article["prism:publicationName"],
    "container-volume": article["prism:volume"],
    "container-issue": article["prism:issue"],
  };

  if (typeof item.authors == "string") item.authors = [item.authors];

  return item;
}

function normaliseMendeley(data){
  Logger.log(data.toSource());

  // TODO

  return {
    "title": data.title,
    "authors": data.authors,
    "date": data.date,
    "doi": data.doi
  }
}

function normaliseGeneric(data){
  Logger.log(data.toSource());
  // TODO
  return data;
}

/** Local spreadshet **/

function loadItems(){
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  var rows = sheet.getLastRow() - 1;

  if (rows < 1) return false;

  var range = sheet.getRange(2, 1, rows, 2);
  var values = range.getValues();

  var items = {};
  for (var i = 0; i < values.length; i++){
    var row = values[i];
    var key = row[0];
    var data = row[1];
    if (key && data){
      //var resultKey = Utilities.base64Encode(key);
      var resultKey = Base64.encode(key);
      items[resultKey] = Utilities.jsonParse(data);
    }
  }

  return items;
}

function saveItems(items){
  var data = [];
  for (var key in items){
    if (typeof known[key] != "undefined") continue;

    var item = items[key];
    //var resultKey = Utilities.base64Decode(key);
    var resultKey = Base64.decode(key);
    data.push([resultKey, Utilities.jsonStringify(item.data), updated, item.data["title"], item.data["date"], item.data["doi"]]);
  }

  if (!data.length) return false;
  
  Logger.log(updated);
  
  Logger.log(data.toSource());
  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  
  var lastRow = Math.max(1, sheet.getLastRow());
  sheet.insertRowsAfter(lastRow, data.length);
  
  var range = sheet.getRange(lastRow + 1, 1, data.length, 6);
  range.setValues(data);
}

/** OAuth **/

function getConsumerKey() {
  var key = ScriptProperties.getProperty("consumerKey");
  return key == null ? "" : key;
}

function getConsumerSecret() {
  var secret = ScriptProperties.getProperty("consumerSecret");
  return secret == null ? "" : secret;
}

function isConfigured() {
  return getConsumerKey() != "" && getConsumerSecret != "";
}

function setupOAuth(){
  if (!isConfigured()) configure(true);

  var oauthConfig = UrlFetchApp.addOAuthService("mendeley");
  oauthConfig.setAccessTokenUrl("http://www.mendeley.com/oauth/access_token/");
  oauthConfig.setRequestTokenUrl("http://www.mendeley.com/oauth/request_token/");
  oauthConfig.setAuthorizationUrl("http://www.mendeley.com/oauth/authorize/");
  oauthConfig.setConsumerKey(getConsumerKey());
  oauthConfig.setConsumerSecret(getConsumerSecret());
}

function authorizeMendeley() {
  setupOAuth();

  UrlFetchApp.fetch("http://api.mendeley.com/oapi/profiles/info/me/",{
      "method": "GET",
      "oAuthServiceName": "mendeley",
      "oAuthUseToken": "always"
  });
}

function configure(auth) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle("Configure Mendeley OAuth");

  var listPanel = app.createGrid(4, 2);
  listPanel.setStyleAttribute("margin-top", "10px")
  listPanel.setWidth("100%");

  var consumerKeyLabel = app.createLabel("Mendeley OAuth Consumer Key:");
  listPanel.setWidget(1, 0, consumerKeyLabel);

  var consumerKey = app.createTextBox();
  consumerKey.setName("consumerKey");
  consumerKey.setWidth("100%");
  consumerKey.setText(getConsumerKey());
  listPanel.setWidget(1, 1, consumerKey);

  var consumerSecretLabel = app.createLabel("Mendeley OAuth Consumer Secret:");
  listPanel.setWidget(2, 0, consumerSecretLabel);

  var consumerSecret = app.createTextBox();
  consumerSecret.setName("consumerSecret");
  consumerSecret.setWidth("100%");
  consumerSecret.setText(getConsumerSecret());
  listPanel.setWidget(2, 1, consumerSecret);

  var saveHandler = app.createServerClickHandler("saveConfiguration");
  saveHandler.addCallbackElement(listPanel);
  var saveButton = app.createButton("Save Configuration", saveHandler);

  var dialogPanel = app.createFlowPanel();
  dialogPanel.add(listPanel);
  dialogPanel.add(saveButton);
  app.add(dialogPanel);
  doc.show(app);
  
  if (auth) throw "Keys are required for OAuth requests";
}

function saveConfiguration(e) {
  ScriptProperties.setProperty("consumerKey", e.parameter.consumerKey);
  ScriptProperties.setProperty("consumerSecret", e.parameter.consumerSecret);
  ScriptProperties.setProperty("serviceUrl", e.parameter.serviceUrl);
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

/** Base64 to/from strings **/

// http://www.webtoolkit.info/javascript-base64.html

var Base64 = {
  // private property
  _keyStr : "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",

  // public method for encoding
  encode : function (input) {
    var output = "";
    var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
    var i = 0;

    input = Base64._utf8_encode(input);

    while (i < input.length) {

      chr1 = input.charCodeAt(i++);
      chr2 = input.charCodeAt(i++);
      chr3 = input.charCodeAt(i++);

      enc1 = chr1 >> 2;
      enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
      enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
      enc4 = chr3 & 63;

      if (isNaN(chr2)) {
        enc3 = enc4 = 64;
      } else if (isNaN(chr3)) {
        enc4 = 64;
      }

      output = output +
      this._keyStr.charAt(enc1) + this._keyStr.charAt(enc2) +
      this._keyStr.charAt(enc3) + this._keyStr.charAt(enc4);

    }

    return output;
  },

  // public method for decoding
  decode : function (input) {
    var output = "";
    var chr1, chr2, chr3;
    var enc1, enc2, enc3, enc4;
    var i = 0;

    input = input.replace(/[^A-Za-z0-9\+\/\=]/g, "");

    while (i < input.length) {

      enc1 = this._keyStr.indexOf(input.charAt(i++));
      enc2 = this._keyStr.indexOf(input.charAt(i++));
      enc3 = this._keyStr.indexOf(input.charAt(i++));
      enc4 = this._keyStr.indexOf(input.charAt(i++));

      chr1 = (enc1 << 2) | (enc2 >> 4);
      chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
      chr3 = ((enc3 & 3) << 6) | enc4;

      output = output + String.fromCharCode(chr1);

      if (enc3 != 64) {
        output = output + String.fromCharCode(chr2);
      }
      if (enc4 != 64) {
        output = output + String.fromCharCode(chr3);
      }

    }

    output = Base64._utf8_decode(output);

    return output;

  },

  // private method for UTF-8 encoding
  _utf8_encode : function (string) {
    string = string.replace(/\r\n/g,"\n");
    var utftext = "";

    for (var n = 0; n < string.length; n++) {

      var c = string.charCodeAt(n);

      if (c < 128) {
        utftext += String.fromCharCode(c);
      }
      else if((c > 127) && (c < 2048)) {
        utftext += String.fromCharCode((c >> 6) | 192);
        utftext += String.fromCharCode((c & 63) | 128);
      }
      else {
        utftext += String.fromCharCode((c >> 12) | 224);
        utftext += String.fromCharCode(((c >> 6) & 63) | 128);
        utftext += String.fromCharCode((c & 63) | 128);
      }

    }

    return utftext;
  },

  // private method for UTF-8 decoding
  _utf8_decode : function (utftext) {
    var string = "";
    var i = 0;
    var c = c1 = c2 = 0;

    while ( i < utftext.length ) {

      c = utftext.charCodeAt(i);

      if (c < 128) {
        string += String.fromCharCode(c);
        i++;
      }
      else if((c > 191) && (c < 224)) {
        c2 = utftext.charCodeAt(i+1);
        string += String.fromCharCode(((c & 31) << 6) | (c2 & 63));
        i += 2;
      }
      else {
        c2 = utftext.charCodeAt(i+1);
        c3 = utftext.charCodeAt(i+2);
        string += String.fromCharCode(((c & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
        i += 3;
      }

    }

    return string;
  }
}
