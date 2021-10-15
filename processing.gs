/*******************************
 TABLE OF CONTENT
   -  processData()
   -  checkInitialized(dataSource, dataSourceSheet)
   -  checkFilterValues(notProcessed)
   -  updateSettings(cValues, row, col, nRows, nCols)
   -  routeData(row)
   -  
   -  
   -  
   -  
   

********************************/

// GLOBAL VARIABLES
let dest_ss, settings, u;



function processData() {
  dest_ss = SpreadsheetApp.getActiveSpreadsheet();
  settings = dest_ss.getSheetByName('Automation').getDataRange().getValues();
  getReferences();
  if (u.power) {
    var test = settings[u.sheetName_loc.r - 1][u.sheetName_loc.c - 1]
    var test1 = settings[u.dataSourceURL_loc.r - 1][u.dataSourceURL_loc.c - 1]
    var dataSourceSheet = SpreadsheetApp.openByUrl(settings[u.dataSourceURL_loc.r - 1][u.dataSourceURL_loc.c - 1]).getSheetByName(settings[u.sheetName_loc.r - 1][u.sheetName_loc.c - 1]);
    var dataSource = dataSourceSheet.getDataRange();
    var lastDataColumn = dataSource.getLastColumn();
    checkInitialized(dataSource, dataSourceSheet);
    var dataSource = dataSource.offset(settings[u.dataHeaderRows_loc.r - 1][u.dataHeaderRows_loc.c - 1], 0, dataSourceSheet.getLastRow());
    var e = dataSource.getValues();
    var notProcessed = e.filter((row) => row[lastDataColumn - 1] != "Processed")
    checkFilterValues(notProcessed);  //adds any new values found in the filter column to the automation tab for user processing later.
    var filterLogic = settings.filter((r, i) => r[u.filterValue_loc.c - 1] != "" && i >= (u.filterValue_loc.r - 1)).map(r => [r[u.filterValue_loc.c - 1], r[u.filterSheetName_loc.c - 1]]);


    e.forEach((row, i) => {
      if (row[lastDataColumn - 1] != "Processed" && row[u.filterColumn - 1] != "") {
        var filterValue = row[u.filterColumn - 1]  
        var sheetName = filterLogic.find(v => v[0] == filterValue)
        if (sheetName == undefined || sheetName[1] == "") {
          sheetName = "Other" 
        } else {
          sheetName = sheetName[1]
        }
        
        var transferedData = routeData(row)

        var dest_sheet = dest_ss.getSheetByName(sheetName);
        dest_sheet.insertRowAfter(dest_sheet.getLastRow());
        dest_sheet.getRange(dest_sheet.getLastRow() + 1, 1, 1, transferedData.length).setValues([transferedData])
        var test = (i + 1) + settings[u.dataHeaderRows_loc.r - 1][u.dataHeaderRows_loc.c - 1]
        dataSourceSheet.getRange((i + 1) + settings[u.dataHeaderRows_loc.r - 1][u.dataHeaderRows_loc.c - 1], lastDataColumn).setValue("Processed");
      }
    })



    var test = 1;
  } else {
    console.log("Power is off")
  }
}



function checkInitialized(dataSource, dataSourceSheet) {
  if (dataSource.getValues()[0][dataSource.getLastColumn() - 1] != "Processed") {
    dataSourceSheet.getRange(1, dataSource.getLastColumn() + 1).setValue("Processed");
    console.log("Initialized data source by adding a 'Processed' column.")
  }
}


function checkFilterValues(notProcessed) {
  var cValues = settings.filter((row, i) => row[u.filterValue_loc.c - 1] != "" && i >= (u.filterValue_loc.r - 1)).map(r => r[u.filterValue_loc.c - 1])
  var nValues = notProcessed.map(r => r[u.filterColumn - 1]).filter((r, i, s) => cValues.indexOf(r) == -1 && s.indexOf(r) === i && r != "")

  if (nValues.length > 0) {
    cValues = cValues.concat(nValues)
    updateSettings(cValues.map(r=> [r]), u.filterValue_loc.r, u.filterValue_loc.c, cValues.length)
  }
  return cValues
}



function updateSettings(cValues, row, col, nRows, nCols) {
  dest_ss.getSheetByName('Automation').getRange(row, col, nRows || 1, nCols || 1).setValues(cValues)
  SpreadsheetApp.flush()
}



function routeData(row) {
  var routing = settings.map(r => [r[u.sourceColumn_loc.c - 1], r[u.toColumn_loc.c - 1], r[u.findSubstring_loc.c -1], r[u.startString_loc.c - 1], r[u.endString_loc.c - 1]]).filter(r => r[0] != "" && r[1] != ""); routing.shift()
  routing = routing.map(r => [SStoNumColumn(r[0]), SStoNumColumn(r[1]), r[2], r[3], r[4]])
  //routing.forEach((v, i, t) => v.forEach((v, i, t) => t[i] = SStoNumColumn(v)))
  var maxCol = findLargest(routing.map(r => r[1]))

  var data = new Array(maxCol)
  routing.forEach((route) => {
    if (route[2]) {
      data[route[1] - 1] = getSubString(row[route[0] - 1], route[3], route[4])
    } else {
      data[route[1] - 1] = row[route[0] - 1]
    }
  })
  return data
}





function test_dataRoute() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Test')
  var e = sheet.getDataRange().offset(1, 0).getValues();

  var data = e.map((row) => {
    if (row[1]) {
      return [getSubString(row[0], row[2], row[3])]
    } else {
      return [row[0]]
    }
  })

  sheet.getRange(2, 5, data.length).setValues(data)
}
