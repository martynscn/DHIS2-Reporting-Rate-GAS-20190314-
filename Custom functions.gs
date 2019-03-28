function updateDataWeekly(sourceSheetName,destSheetName,sourceSheetStartRow,sourceSheetStartCol,destStartCol, temp_destStartRow) {
  
  var sourceSheetName = sourceSheetName || 'extract_data';
  var destSheetName = destSheetName || 'IncrementalData';
  var sourceSheetStartRow = sourceSheetStartRow || 2;
  var sourceSheetStartCol = sourceSheetStartCol || 1;
  var destStartCol = destStartCol || 1;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if(!ss.getSheetByName(destSheetName)) {
    ss.insertSheet(destSheetName);
  }

  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var lr_sourceSheet = sourceSheet.getLastRow();
  var lc_sourceSheet = sourceSheet.getLastColumn();
  var destSheet = ss.getSheetByName(destSheetName);

  var sourceData = sourceSheet.getRange(sourceSheetStartRow, sourceSheetStartCol, lr_sourceSheet-1, lc_sourceSheet).getValues();
  var lr_destSheet = destSheet.getLastRow();
  var destStartRow = Math.max(temp_destStartRow || 0, (lr_destSheet) + 1);
  destSheet.getRange(destStartRow, destStartCol, sourceData.length, sourceData[0].length).setValues(sourceData);
}

function updateDataFrequently_fixedRange(sourceSheetName,destSheetName,sourceSheetStartRow,sourceSheetStartCol,sourceSheetNumRows, sourceSheetNumCols, destSheetStartRow,destSheetStartCol) {
  
  var sourceSheetName = sourceSheetName || 'extract_data';
  var destSheetName = destSheetName || 'IncrementalData';
  var sourceSheetStartRow = sourceSheetStartRow || 2;
  var sourceSheetStartCol = sourceSheetStartCol || 1;
  var sourceSheetNumRows = sourceSheetNumRows || 3;
  var sourceSheetNumCols = sourceSheetNumCols || 71;
  var destSheetStartRow = destSheetStartRow || 2;
  var destSheetStartCol = destSheetStartCol || 1;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if(!ss.getSheetByName(destSheetName)) {
    ss.insertSheet(destSheetName);
  }

  var sourceData = ss.getSheetByName(sourceSheetName).getRange(sourceSheetStartRow, sourceSheetStartCol, sourceSheetNumRows, sourceSheetNumCols).getValues();
  ss.getSheetByName(destSheetName).getRange(destSheetStartRow, destSheetStartCol, sourceData.length, sourceData[0].length).setValues(sourceData);
//  var sourceRange = ss.getSheetByName(sourceSheetName).getRange(sourceSheetStartRow, sourceSheetStartCol, sourceSheetNumRows, sourceSheetNumCols);
//  var destRange = ss.getSheetByName(destSheetName).getRange(destSheetStartRow, destSheetStartCol, sourceSheetNumRows, sourceSheetNumCols);
//  sourceRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
//  sourceRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
//  sourceRange.copyFormatToRange( ss.getSheetByName(destSheetName), destSheetStartCol, destSheetStartCol + sourceSheetNumCols , destSheetStartRow, destSheetStartRow + sourceSheetNumRows);

  
}


function transpose(a) {
  return Object.keys(a[0]).map(function (c) {
    return a.map(function (r) {
        return r[c];
    });
  });
}

function myVlookup(sourceArray, tableToLookup,columnToSearch,columnToReturn) {
  var o = [];
  var columnToSearch = Number(columnToSearch - 1);
  var columnToReturn = Number(columnToReturn - 1);
  for (var i = 0; i < sourceArray.length; i++) {
    for (var j = 0; j < tableToLookup.length; j++) {
      if(sourceArray[i] == tableToLookup[j][columnToSearch]) {
        o.push(tableToLookup[j][columnToReturn]);
        break;
      }
    }
  }
  return o;
}

function sendSlackMessage(subject, body, postUrl) {
  var postUrl = postUrl || "https://hooks.slack.com/services/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";
  var subject = subject || "No subject";
  var body = body || "No body";
  var message = subject + '\n\n' + body;
  var jsonData =
  {
     "text" : message
  };
  var payload = JSON.stringify(jsonData);
  var options =
  {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : payload
  };

  UrlFetchApp.fetch(postUrl, options);
}