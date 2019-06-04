/*
Copy this function over the existing Code.gs file in a Google Sheet
Save the Sheet/Script
Publish App 
- Execute as You
- Avaiable to Anyone (including anonymous)
Authorize App when prompted
The URL will act as the Webhook URL for Meraki to send alerts
Configure Trigger to run function "macro~" every minute
*/

// Settings - Modify this with your values
// *************************

// User Defined in the Script
var API_KEY = '';
var BASE_URL = 'https://api.meraki.com/api/v0';
var ORG_ID = '';
var NET_ID = '';
var TIMESPAN = '';

// User Defined in a Sheet
var SHEET_NAME = "settings"
var API_KEY_SHEET_CELL = "B3";
var API_KEY_SHEET_CELL_LABEL = "A3";
var BASE_URL_SHEET_CELL = "B4";
var BASE_URL_SHEET_CELL_LABEL = "A4";

var TIMESPAN_SHEET_CELL = "B6";
var TIMESPAN_SHEET_CELL_LABEL = "A6";

var macroQueueSheetTitle = 'Queue: VPN-Reboot';

// *************************
// Initialize Settings Sheet and Environment Variables

// find or create settings sheet
var ss = SpreadsheetApp.getActiveSpreadsheet();
if (ss.getSheetByName(SHEET_NAME) == null){
  ss.insertSheet(SHEET_NAME); 
  ss.getRange(API_KEY_SHEET_CELL_LABEL).setValue('API KEY:');
  ss.getRange(BASE_URL_SHEET_CELL_LABEL).setValue('BASE_URL:');

  ss.getRange(TIMESPAN_SHEET_CELL_LABEL).setValue('REBOOT Device after VPN offline after X seconds: ');
  ss.getRange(TIMESPAN_SHEET_CELL).setValue(180);
}
var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

// assign settings
var settings = {};
settings.apiKey = settingsSheet.getRange(API_KEY_SHEET_CELL).getValue() || API_KEY; 
settings.baseUrl = settingsSheet.getRange(BASE_URL_SHEET_CELL).getValue() || BASE_URL;
settings.timespan = settingsSheet.getRange(TIMESPAN_SHEET_CELL).getValue() || TIMESPAN;

/**
 Utility Functions
*/

function formatTime(merakiTimeStamp){
  // removes milliseconds, since GAS version of JS does not parse ISO properly
  var splitTime = merakiTimeStamp.split(":");
  var splitSeconds = splitTime[2].split('.') ;
  return new Date(splitTime[0]+":"+splitTime[1]+":"+splitSeconds[0]);
}

// Flattens a nested object for easier use with a spreadsheet
function flattenObject(ob) {
   var toReturn = {};	
	for (var i in ob) {
		if (!ob.hasOwnProperty(i)) continue;		
		if ((typeof ob[i]) == 'object') {
			var flatObject = flattenObject(ob[i]);
			for (var x in flatObject) {
				if (!flatObject.hasOwnProperty(x)) continue;				
				toReturn[i + '.' + x] = flatObject[x];
			}
		} else {
			toReturn[i] = ob[i];
		}
	}
	return toReturn;
};

// formats a key/value to UTC time based on selected keys
function changeTimeFormat(key,value){
  var keysToFormat = ['sentAt', 'occurredAt','alertData.timestamp'];
  if(keysToFormat.indexOf(key) > -1){
        var date = new Date(value*1000).toUTCString();
        return date 
  }else{
    return value;
  }
}
// 
function searchSheet(sheet, searchString){
  if(!searchString){return}
  if(!sheet){return}
  
  
  var textFinder = sheet.createTextFinder(searchString)
  var textRange = textFinder.findNext();
  Logger.log('searchSheet textFinder.findNext(): '+ textFinder.findNext());
  if(!textRange){return}
  var results = {
    row: textRange.getRow(),
    column: textRange.getColumn(),
    rowData: sheet.getRange(textRange.getRow(), 1, 1, sheet.getLastColumn()).getValues(),
    columnData: sheet.getRange(1, textRange.getColumn(), sheet.getLastRow(), 1).getValues() 
  }

  Logger.log(results);
  return results;
}

/**
Display to Sheets Utilties
*/

function setHeaders(sheet, values){
   var headerRow = sheet.getRange(1, 1, 1, values.length)
    headerRow.setValues([values]);  
    headerRow.setFontWeight("bold").setHorizontalAlignment("center");
}

function display(data, sheetTitle){
  
  // Flatten JSON object and extract keys and values into seperate arrays
  var flat = flattenObject(data);
  var keys = Object.keys(flat);
  var values = [];
  var headers = [];
  var alertSheet;
  
  // Find or create sheet for alert type and set headers
  var alertType = sheetTitle || data['alertType'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(alertType) == null){
    ss.insertSheet(alertType); 
    // Create Headers and Format
    alertSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(alertType);
    alertSheet.setColumnWidths(1, keys.length, 200)
    headers = keys;
  }else {
    alertSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(alertType);
    // retrieve existing headers
    headers = alertSheet.getRange(1, 1, 1, alertSheet.getLastColumn() || 1).getValues()[0]; 
   
    // add any additional headers
    var newHeaders = [];
    newHeaders = keys.filter(function(k){ return headers.indexOf(k)>-1?false:k;});
    newHeaders.forEach(function(h){
      headers.push(h);
    });  
  }
  Logger.log('headers: ' + headers);
  setHeaders(alertSheet, headers);
  
  // push values based on headers
  headers.forEach(function(h){
    values.push(flat[h]);
  });
  
  // Insert Data into Sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(alertType);
  var lastRow = Math.max(sheet.getLastRow(),1);
  sheet.insertRowAfter(lastRow); 
  sheet.getRange(lastRow + 1, 1, 1, headers.length).setValues([values]).setFontWeight("normal").setHorizontalAlignment("center");
}
  
// Webhook GET request. Simply verifies that server is reachable.
function doGet(e) {
  return HtmlService.createHtmlOutput("Meraki Webhook Google Sheets");
}

// Webhook Receiver - triggered with post to pusblished App URL.
function doPost(e) {
  var params = JSON.stringify(e.postData.contents);
  params = JSON.parse(params);
  var postData = JSON.parse(params);
  
  display(postData);
  workflow(postData);
  
 // HTTP Response
 return ContentService.createTextOutput("post request received");
}

function removeAllRecords(sheet, searchTerm){
  while(true){
    var searchResults = searchSheet(sheet,searchTerm)  
    if(!searchResults){return}
    var rowIndex = searchResults.row;
    sheet.deleteRow(rowIndex)
  }
}

/**
Workflow to trigger actions based on alertType
*/

function workflow(postData){
  switch (postData["alertType"]){
      case "VPN connectivity changed":
      Logger.log('workflow - VPN connectivity changed for `networkId`: ' + postData["networkId"]);
      // ** Reboot device if VPN has resumed connectivity (hack to reset AutoVPN)
      // VPN goes down and does not connect after 3 minutes -- TODO --> 
      
      // store False event in temp sheet,
      // wait for True event, 
      // if True event has not been seen in X seconds, --> reboot
      // if True event has been seen, clear False records 
         
      if(!postData.alertData.connectivity || postData.alertData.connectivity === "false"){
        // Store alert in temp sheet  
        display(postData, macroQueueSheetTitle)
      }else{
        // Remove alert from temp sheet
        Logger.log('VPN connectivity resumed, cancelling reboot');
         
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(macroQueueSheetTitle);  
        removeAllRecords(sheet, postData["deviceSerial"])
      }  
      
      break;
  }    
}

function getSheetValues(sheet){
  var data = sheet.getDataRange().getValues();
  return data;
}


/**
* Macro to run periodically from scheduled Trigger
*/
function macroRebootDevicesFailedVPN(){
  /*
  read sheet 'Queue: VPN-Reboot'
  get column number for occurredAt time
  get column value for each row 
  if value is greater than diftime
  reboot device
  log record in sheet "Workflow: Logs"
  delete temp record row from Queue 
  */
  
  // Define Sheet to temporarily store recent alerts
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(macroQueueSheetTitle); 
  var sheetData = getSheetValues(sheet);
  Logger.log("sheetData: " + sheetData);
  
  // Find the timestamps and table location for each alert
  var searchResults = searchSheet(sheet, "occurredAt");
  if(!searchResults){return}
  var rowIndex = searchResults.row + 1;
  var occurredAtIndex = searchResults.column - 1;
  Logger.log("occurredAtIndex: "+ (occurredAtIndex));
  var deviceSerialIndex = searchResults.rowData[0].indexOf("deviceSerial");
  var networkIdIndex = searchResults.rowData[0].indexOf("networkId");
  Logger.log("deviceSerialIndex: "+ (deviceSerialIndex));
  
  var rebootedDevices = [];
  // Iterate through each record
  const now = new Date();
  for (var i = 1; i < sheetData.length; i++) {   
    var occurredAt = sheetData[i][(occurredAtIndex)]
    occurredAt = formatTime(occurredAt);
    
    // Compare time of event and reboot device if time has exceeded
    const timediff = now - occurredAt;
    Logger.log("time difference: " + timediff);
    
    if( timediff > settings.timespan*1000){   
      // reboot device 
      const deviceSerial = sheetData[i][(deviceSerialIndex)];
      const networkId = sheetData[i][(networkIdIndex)]
      Logger.log("rebooting device for network/serial: " + networkId + " : " + deviceSerial);
      
      if(rebootedDevices.indexOf(deviceSerial) < 1){
        Logger.log('deviceSerial: ' + deviceSerial)
        Logger.log('rebootedDevices.indexOf(deviceSerial)' + rebootedDevices.indexOf(deviceSerial));
        const responseJson = rebootDevice(networkId, deviceSerial);
        Logger.log('responseJson: '+ responseJson);
        
        if(responseJson["success"]){
          // Clear Queue and skip future serials
          rebootedDevices.push(deviceSerial);
          removeAllRecords(sheet, deviceSerial)
        }
        
        // report data
        var reportData = {};  
        reportData.workflow = {
          "rebootStatus":responseJson,
          "timestamp": now
        };
        
        // create an object from the headers and the given row values
        var alertData = {};
        var keys = sheetData[0];
        var values = sheetData[i];
        for (i = 0; i < keys.length; i++) {
          alertData[keys[i]] = values[i];
        }
        
        reportData.alertData = alertData;
        display(reportData, "Logs: VPN workflow");
        /*
        sheetData[0].forEach(function(key, index){
          alertData[key] = sheetData[
        }
        reportData.alertData = sheetData[i];
        display(reportData, "Logs: VPN workflow");
        */
      }
    }else{
      Logger.log("not time to reboot for device: " + deviceSerial)
     }
  }    
 
}



// Meraki API handler

function rebootDevice(networkId, serial){
  var options = {
    "headers":{"X-Cisco-Meraki-API-Key": settings.apiKey},
    "method" : "POST",
    "contentType" : "application/json",
    "payload" : ""
  };
  var response = UrlFetchApp.fetch(settings.baseUrl+"/networks/"+networkId+"/devices/"+serial+"/reboot", options);
  var data = response.getContentText();
  Logger.log('rebootDevice data: '+data);
  var json = JSON.parse(data);
  return json;   
}

/**
// Test & Utility Functions
*/

var testData = {
    "sentAt": "2019-01-29T21:39:03.249388Z",
    "alertId": "629378047939324609",
    "version": "0.1",
    "alertData": {
        "vpnType": "site-to-site",
        "peerIdent": "b321882f383b1b1244497c81efbd157f",
        "peerContact": "1.2.3.4:1024",
        "connectivity": false
    },
    "alertType": "VPN connectivity changed",
    "deviceMac": "aa:bb:cc:dd:ee:ff",
    "deviceUrl": "https://n1.meraki.com/.../manage/nodes/new_wired_status",
    "networkId": "L_1234567890",
    "deviceName": "Device",
    "networkUrl": "https://n1.meraki.com/.../manage/nodes/wired_status",
    "occurredAt": "2019-01-29T21:32:00.243000Z",
    "networkName": "Network",
    "deviceSerial": "Q2BX-9QRR-XXXX",
    "sharedSecret": "supersecret",
    "organizationId": "123456",
    "organizationUrl": "https://n1.meraki.com/o/.../manage/organization/overview",
    "organizationName": "Organization"
};



function test(){ 
  display(testData);
}

function testWorkflow(){
  workflow(testData);
}

function testSearchString(){
  searchString("L_1234567890");
};
