function extract_info(metadata_dimensions) {
  try {
    var contentJson = InitiateExtraction();
    var metadata = contentJson.metaData;
    var metadata_dimensions = metadata.dimensions;
    var dx = [];
    var ou = [];
    var pe = [];
    
    for (i in metadata_dimensions.dx) {
      dx.push([metadata_dimensions.dx[i]]);
    }
    for (i in metadata_dimensions.ou) {
      ou.push([metadata_dimensions.ou[i]]);
    }
    for (i in metadata_dimensions.pe) {
      pe.push([metadata_dimensions.pe[i]]);
    }
    var error = false;
  }
  catch(err) {
    var error = true;
    return [['Error in retrieving dx'],['Error in retrieving ou'],['Error in retrieving pe']];
  }
  finally {
    if(error == false) {
      return [dx,ou,pe]; 
    } else if(error = true) {
      return [['Error in retrieving dx'],['Error in retrieving ou'],['Error in retrieving pe']];
    }
  }
  
}
function update_extract_info_in_sheet() {
  var ss = SpreadsheetApp.openById('1Y8ebgjLE0_pRsi_AOyTCXG3cgIDEUQqtBpw_L9ndMDE');
  var dx = extract_info()[0];
  var ou = extract_info()[1];
  var pe = extract_info()[2];
  var firstRow = [['dx','ou','pe']];
  var sheetname = 'extract_info';
  if(!ss.getSheetByName(sheetname)) {
    ss.insertSheet(sheetname);
  }
  ss.getSheetByName(sheetname).getRange(1, 1, firstRow.length, firstRow[0].length).setValues(firstRow);
  ss.getSheetByName(sheetname).getRange(2, 1, dx.length, dx[0].length).setValues(dx);
  ss.getSheetByName(sheetname).getRange(2, 2, ou.length, ou[0].length).setValues(ou);
  ss.getSheetByName(sheetname).getRange(2, 3, pe.length, pe[0].length).setValues(pe);
}

function extract_codes(contentJson) {
    var contentJson = contentJson || InitiateExtraction();
    var metadata = contentJson.metaData;
    var metadata_items = metadata.items;
    var codes = [];
    for (i in metadata_items) {
      codes.push([i,metadata_items[i].name]);
    }
    var firstRow = ['code','code value'];
    codes.unshift(firstRow);
  return codes;
}

function update_codes_in_sheet() {
  var ss = SpreadsheetApp.openById('1Y8ebgjLE0_pRsi_AOyTCXG3cgIDEUQqtBpw_L9ndMDE');
  var contentJson = InitiateExtraction();
  var codes = extract_codes(contentJson);
  var sheetname = 'codes';
  if(!ss.getSheetByName(sheetname)) {
    ss.insertSheet(sheetname);
  }
  ss.getSheetByName(sheetname).getRange(1, 1, codes.length, codes[0].length).setValues(codes);
}
function getResponseandCode(baseUrl,target, uname, pwd) {
  var baseUrl = baseUrl || "https://dhis2nigeria.org.ng/dhis";
  var unameProp = uname || 'nhmis_username';
  var username = ScriptProperties.getProperty(unameProp);
  var pwdProp = pwd || 'nhmis_password';
  var password = ScriptProperties.getProperty(pwdProp);
  var target = target || "/api/29/analytics.json?&dimension=dx:lyVV9bPLlVy.REPORTING_RATE&dimension=pe:LAST_MONTH&dimension=ou:s5DPBsdoE8b;LEVEL-1;LEVEL-2";
  var url = baseUrl + (target == 'none' ? '' : target);
  var slackPostUrl = 'https://hooks.slack.com/services/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'
  var headers = {
        "Authorization": "Basic " + Utilities.base64Encode(username + ':' + password)
      };
    var options =
        {
          "method"  : "GET",
          'headers': headers,
          "muteHttpExceptions": true
        };
  try {
    var response = UrlFetchApp.fetch(url,options);
    sendSlackMessage('Successful Request: ' + baseUrl , 'Accessing url: ' + url + '\nwith username: ' + username + ' was successful',slackPostUrl);
  }
  catch(error) {
     sendSlackMessage('Failed Request: ' + baseUrl , 'Trying to access url: ' + url + '\nwith username: ' + username + ' failed woefully.\nError was ' + error,slackPostUrl);
  }
  var responseCode = response.getResponseCode();
  return [response,responseCode];
}
function InitiateExtraction() {
    //        paste0("&displayProperty=", "NAME")paste0("&outputIdScheme=", "NAME")paste0("&tableLayout=", "true")paste0("&columns=", "dx")paste0("&rows=", "pe;ou")paste0("&skipRounding=", "true")
  var baseUrl = "https://dhis2nigeria.org.ng/dhis";
  var target = "/api/29/analytics.json?&dimension=dx:lyVV9bPLlVy.REPORTING_RATE&dimension=pe:LAST_MONTH&dimension=ou:s5DPBsdoE8b;LEVEL-1;LEVEL-2";
  var uname = 'nhmis_username';
  var pwd = 'nhmis_password';
  
  var response = getResponseandCode(baseUrl,target, uname, pwd)[0];
  var content = response.getContentText();
  var contentJson = JSON.parse(content);
  return contentJson;
}

  SORT_ORDER = [
    {column: 2, ascending: false},  // 3 = column number, sorting by descending order
    {column: 1, ascending: true} // 1 = column number, sort by ascending order 
    ];

/**
* Retrieves the NHMIS reporting rates for last month[later allow it to accept parameters from the API].
*
* //@param {number or string in double quotes e.g. 123456 or "123456"} teamId The team Id.
* 
* @return An array of the NHMIS reporting rates.
* @customfunction
*/
function getNHMISReportingRatesFn() {
  var contentJson = InitiateExtraction();
  //var contentJsonKeys = Object.keys(contentJson); // [headers, metaData, rows, height, width]
  var metadata = contentJson.metaData;
  //var metadataKeys = Object.keys(metadata) //[items, dimensions]
  var contentRows = contentJson.rows;
  var codes = extract_codes(contentJson);
  
  var dx = extract_info(metadata.dimensions)[0];
  var ou = extract_info(metadata.dimensions)[1];
  var pe = extract_info(metadata.dimensions)[2];

  var statesCode = contentRows.map(function (row) {
    return row[2];
  });
  
  var statesName = [myVlookup(statesCode,codes,1,2)];
  var contentRows_tranposed = transpose(contentRows);
  contentRows = contentRows_tranposed.slice(3);
  var completeData = statesName.concat(contentRows);
  


  var cur_period = myVlookup(pe[0],codes,1,2);
  var cur_dx = myVlookup(dx[0],codes,1,2);
  var top_left_rows = [['date Extracted','Period','Indicators'],[new Date(),cur_period[0],cur_dx[0]]];
  

  
  var completeDataT = transpose(completeData).sort();
//  var sortData = completeDataT.sort([{column: 1, ascending: true}, {column: 2, ascending: false}]);
//  Logger.log(completeDataT);
  var top_left_rowsT = transpose(top_left_rows);
  var entireData = transpose(top_left_rowsT.concat(completeDataT));

  return [entireData,cur_period];
}
function updateNHMISReportingRates(ss) {
  var getNHMISReportingRates = getNHMISReportingRatesFn();
  var entireData = getNHMISReportingRates[0];
  var cur_period = getNHMISReportingRates[1];
  var sheetname = 'extract_data';
  updateEntireOnSheet(sheetname,entireData);
    
  var title = 'NHMIS Reporting Rates updates for period: ' + cur_period + ' has been updated on ' + new Date();
  var body = ""; // later format data to be in a table;
  sendSlackMessage(title, body);
}

function updateEntireOnSheet(sheetname,content) {
  var S = SpreadsheetApp.getActiveSpreadsheet();
  if(!S.getSheetByName(sheetname)) {
    S.insertSheet(sheetname);
  }
  var ss = S.getSheetByName(sheetname);
  ss.getRange(1, 1, content.length, content[0].length).setValues(content);
  var lastRow = ss.getLastRow();
  var lastCol = ss.getLastColumn();
  var maxRow = ss.getMaxRows();
  var maxCol = ss.getMaxColumns();
  var trimGap = 1;
  try {
  ss.deleteRows(lastRow + trimGap, maxRow - (lastRow + trimGap));
  ss.deleteColumns(lastCol + trimGap, maxCol - (lastCol + trimGap));
  }
  catch (err) {
    Logger.log('The rows or columns were out of boounds or this is the exact error' + err);
  }
}