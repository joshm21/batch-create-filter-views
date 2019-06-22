function doGet(e) {
  
  if (e.queryString && 'jsonpCallback' in e.parameter){
    var cbFnName = e.parameter['jsonpCallback'];
    var scriptText = "window." + cbFnName + "();";
    return ContentService.createTextOutput(scriptText).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  else if (e.queryString && ('auth' in e.parameter || 'redirect' in e.parameter)){
    var rawHtml = '<p>You have successfully logged in! Please close this tab and refresh the previous page.</p>';
    if ('redirect' in e.parameter){
      rawHtml += '<br/><a href="' + e.parameter['redirect'] + '">Previous Page</a>';
    }
    return HtmlService.createHtmlOutput(rawHtml);
  }
  else {
    var template = HtmlService.createTemplateFromFile("index");
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1, shrink-to-fit=no")
      .setTitle("Batch Create Filter Views");
  }
}

function getSheetNames(spreadsheetUrl) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var sheets = spreadsheet.getSheets();
  var sheetNames = sheets.map(function(sheet) {return sheet.getName()});
  return sheetNames;
}

function getColumnNames(spreadsheetUrl, sheetName, headerRow) {
  var columnNames = SpreadsheetApp.openByUrl(spreadsheetUrl).getSheetByName(sheetName).getDataRange().getValues()[headerRow - 1];
  return JSON.stringify(columnNames);
}

function createFilterViews(spreadsheetUrl, sheetName, headerRow, columnNames, filterNames, sortRules, filterRules) {
  var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var spreadsheetId = spreadsheet.getId();
  var sheet = spreadsheet.getSheetByName(sheetName);
  var sheetId = sheet.getSheetId();
  var requests = [];
  for (var i = 0; i < filterNames.length; i++) {
    var title = filterNames[i];
    var range = {
      sheetId: sheetId,
      startRowIndex: headerRow - 1,
      startColumnIndex: 0
    }
    var sortSpecs = [];
    for (var j = 0; j < sortRules.length; j++) {
      var sortRule = sortRules[j];
      var sortSpec = {
        dimensionIndex: columnNames.indexOf(sortRule.column),
        sortOrder: sortRule.order == "Ascending" ? "ASCENDING": "DESCENDING"
      }
      sortSpecs.push(sortSpec);
    }
    var criteria = {}
    for (var column in filterRules) {
      var filterRule = filterRules[column];
      var columnIndex = columnNames.indexOf(filterRule.column);
      criteria[columnIndex] = {
        condition: {
          type: filterRule.type,
          values: {
            userEnteredValue: filterRule.value == "[Filter Name]" ? title : filterRule.value
          }
        }
      }
    }
    var filterViewRequest = {
      addFilterView: {
        filter: {
          title: title,
          range: range,
          sortSpecs: sortSpecs,
          criteria: criteria
        }
      }
    }
    requests.push(filterViewRequest);
  }

  var resource = { requests: requests };
  var response = Sheets.Spreadsheets.batchUpdate(resource, spreadsheetId);
  logCreateFilterViews_(response, sheetName, sortRules, filterRules);
}

function deleteFilterViews(spreadsheetUrl, sheetName) {
  var spreadsheetId = SpreadsheetApp.openByUrl(spreadsheetUrl).getId();
  var filterViews = Sheets.Spreadsheets.get(spreadsheetId, { ranges: sheetName, fields: "sheets/filterViews"}).sheets[0].filterViews;
  if (filterViews == undefined) { return }
  var requests = filterViews.map(function(filterView) {
    return { deleteFilterView: {filterId: filterView.filterViewId} }
  });
  var resource = { requests: requests };
  Sheets.Spreadsheets.batchUpdate(resource, spreadsheetId);
}

function logCreateFilterViews_(response, sheetName, sortRules, filterRules) {
  var timestamp = new Date();
  var spreadsheet = SpreadsheetApp.openById(response.spreadsheetId);
  var spreadsheetUrl = spreadsheet.getUrl();
  var logSheet = spreadsheet.getSheetByName("Filter View Logs");
  if (logSheet == null) {
    var logSheet = spreadsheet.insertSheet("Filter View Logs");
    var headers = ["Name", "Link", "Range", "Sort Rules", "Filter Rules", "Id", "Create Date"];
    logSheet.deleteColumns(headers.length + 1, 19)
    logSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setValues([headers]);
  }
  var data = response.replies.map(function(singleResponse) {
    var filter = singleResponse.addFilterView.filter;
    var name = filter.title;
    var link = spreadsheetUrl + "#gid=" + filter.range.sheetId.toString() + "&fvid=" + filter.filterViewId.toString();
    var range = sheetName + "!" + toColumnLetter_(filter.range.startColumnIndex + 1) + (filter.range.startRowIndex + 1) + ":" + toColumnLetter_(filter.range.endColumnIndex) + filter.range.endRowIndex;   
    var sheetSortRules = sortRules.map(function(rule) {
      return rule.column + " (" + rule.order + ")";
    }).join(", then by ");
    var filters = [];
    for (var prop in filterRules) {
      var rule = filterRules[prop];
      var val = rule.value == "[Filter Name]" ? name : rule.value
      filters.push(rule.column + " " + rule.readable + " " + val);
    }
    var sheetFilterRules = filters.join("; ");
    var id = filter.filterViewId;
    var createDate = timestamp;
    var row = [name, link, range, sheetSortRules, sheetFilterRules, id, createDate]
    return row;
  });
  logSheet.getRange(logSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data)
}

function toColumnLetter_(columnNumber) {
  for (var ret = '', a = 1, b = 26; (columnNumber -= a) >= 0; a = b, b *= 26) {
    ret = String.fromCharCode(parseInt((columnNumber % b) / a) + 65) + ret;
  }
  return ret;
}
