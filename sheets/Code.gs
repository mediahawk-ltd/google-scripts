/** --- INSTRUCTIONS ---
 * 1. Edit the values in the CONFIG section
 * 2. Execute the main() method
 */

/** --- CONFIG --- **/
var USER = ""; //Mediahawk API username
var PASSWORD = ""; //Mediahawk API password
var SPREADSHEET_URL = ""; //Full URL of Google Spreadsheet to insert data into
var SHEET_NAME = ""; //Sheet name
var API_FIELDS = [""]; //Array of Mediahawk API field names to insert into Google Spreadsheet in order
/** --- END CONFIG --- **/

/**
 * Gets yesterday's calls for the specified account from the Mediahawk API and updates the selected
 * spreadsheet with the fields that have been specified
 */
function main() {
  var xml = fetchCallsFromAPI(USER, PASSWORD);
  var calls = parseXmlForCalls(xml, API_FIELDS);
  if(calls.length > 0) {
    insertIntoSpreadsheet(SPREADSHEET_URL, SHEET_NAME, calls);
  }
}

/**
 * Collect yesterday's calls from the api and return the XMLService
 * @param {string} username The username for the API
 * @param {string} password The password for the API
 * @return {Document} The XML document returned from the API
 */
function fetchCallsFromAPI(username, password) {
  var authHash = Utilities.base64Encode(username + ":" + password);
  var currentDate = Utilities.formatDate(getYesterday(), "GMT", "dd/MM/yyyy");
  var xml = "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><SOAP-ENV:Envelope xmlns:SOAP-ENV=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:SOAP-ENC=\"http://schemas.xmlsoap.org/soap/encoding/\" xmlns:tns=\"https://www.reports.mediahawk.co.uk/api/v1_9/\"><SOAP-ENV:Body><tns:GetCallsByDate xmlns:tns=\"https://www.reports.mediahawk.co.uk/api/v1_9/\"><CallDate>" + currentDate + "</CallDate></tns:GetCallsByDate></SOAP-ENV:Body></SOAP-ENV:Envelope>";
  var options =
  {
    "method" : "post",
    "contentType" : "text/xml",
    "payload" : xml,
    "headers" : {
      "Authorization" : "Basic " + authHash,
      "SOAPAction" : "https://www.reports.mediahawk.co.uk/api/v1_9/#GetCallsByDate"
    }
  };
  var result = UrlFetchApp.fetch("https://www.reports.mediahawk.co.uk/api/v1_9/callanalyticsservice.php", options);
  return XmlService.parse(result);
}

/**
 * Process the XML returned from the web service and return an array of calls with the data field
 * specified in the config
 * @param {Document} xml The xml document to parse
 * @param {array} fields The data fields we want from the API data
 * @return {array} Calls containing only data for fields specified
 */
function parseXmlForCalls(xml, fields) {
  var soap = XmlService.getNamespace('http://schemas.xmlsoap.org/soap/envelope/');
  var mh = XmlService.getNamespace('https://www.reports.mediahawk.co.uk/api/v1_9/');
  var calls = xml.getRootElement().getChild('Body', soap).getChild('GetCallsByDateResponse', mh).getChild('CallData').getChildren('item');
  var callsForSpreadsheet = [];
  for (var i = 0; i < calls.length; i++) {
    callsForSpreadsheet.push(getCallValues(calls[i], fields));
  }
  return callsForSpreadsheet;
}


/**
 * Take a call with all data fields and return only the specific ones requested
 * @param {Element} call The call XML element
 * @param {array} fields The data fields we want from the API data
 * @return {array} A call with just the data fields requested
 */
function getCallValues(call, fields) {
  var callValues = [];
  for(var i = 0; i < fields.length; i++) {
    callValues.push(call.getChild(fields[i]).getText());
  }
  return callValues;
}

/**
 * Insert out data into a spreadsheet at the end of the rows
 * All calls are assumed to have the same number of data fields and we will start inserting data
 * at column A
 * @param {string} spreadsheetUrl The URL of the spreadsheet to insert into
 * @param {string} sheetName The name of the sheet to insert data into (will be created if doesn't exist)
 * @param {array} data The data to insert
 */
function insertIntoSpreadsheet(spreadsheetUrl, sheetName, data) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if(!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, data.length, data[0].length).setValues(data);
}

/**
 * Return a date object from 24 hours ago
 * @return {Date} Yesterday's date
 */
function getYesterday() {
  var milliseconds = 1000 * 60 * 60 * 24;
  var now = new Date();
  return new Date(now.getTime() - milliseconds);
}