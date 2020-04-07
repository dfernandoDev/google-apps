// https://developers.google.com/apps-script/reference/spreadsheet/

function removeLastYearSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  for (var i = 0; i < sheets.length ; i++ ) {
    var sheet = sheets[i];
    if (sheet.getName() == 'Total' || sheet.getName() == 'CurrentWk') {
      // skip
    }
    else {
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Babysitter');

  menu.addItem('Archieve Current Week', 'ArchieveCurrentWeek')
  menu.addItem('Clean Current Week', 'CleanCurrentWeek')
  menu.addItem('Archieve & Clean Current for Last Week', 'ArchieveCleanLastWeek')
  menu.addToUi();
}

function ArchieveCleanLastWeek() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var cursheet = spreadsheet.getSheets()[0];
  // set the current week to previous weeks
  cursheet.getRange('B19').setValue('1');

  ArchieveCurrentWeek();
  cursheet.activate();
  CleanCurrentWeek();
}

function CleanCurrentWeek() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var basesheet = spreadsheet.getSheets()[0];
  spreadsheet.setActiveSheet(basesheet);

  spreadsheet.getRange('B2').setValue('3:30:00 PM');
  spreadsheet.getRange('B3:B6').activate();
  spreadsheet.getRange('B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C2:C6').activate();
  spreadsheet.getRange('B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  spreadsheet.getRange('E2:E6').clearContent();

  // set pay
  spreadsheet.getRange('B8').setValue('12');
  // reset additional $ & hrs
  spreadsheet.getRange('B9').setValue('0');
  spreadsheet.getRange('B10').setValue('0');

  // set rounding values
  spreadsheet.getRange('H15').setValue('0.5');
  spreadsheet.getRange('H17').setValue('10');

  // reset week
  spreadsheet.getRange('B19').setValue('0');

  // clear paid info
  spreadsheet.getRange('B22').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B23').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  // payment method
  spreadsheet.getRange('B23').setValue('Venmo');

  // TEMP - DI - Wed
  // spreadsheet.getRange('B4').setValue('4:10:00 PM');
  // spreadsheet.getRange('C4').setValue('4:10:00 PM');
  // spreadsheet.getRange('E4').setValue('DI');
  // END - DI

    // TEMP - DI - Fri
  // spreadsheet.getRange('B6').setValue('4:10:00 PM');
  // spreadsheet.getRange('C6').setValue('4:10:00 PM');
  // spreadsheet.getRange('E6').setValue('DI');
  // END - DI
}

function ArchieveCurrentWeek() {
  var basesheet = SpreadsheetApp.getActiveSpreadsheet();
  basesheet.duplicateActiveSheet()
  var basecopy = SpreadsheetApp.getActiveSpreadsheet();
  basecopy.renameActiveSheet(basecopy.getRange('C19').activate().getValue());
  // reset time calculation
  basecopy.getRange('D2').setValue(basecopy.getRange('D2').getValue());
  basecopy.getRange('D3').setValue(basecopy.getRange('D3').getValue());
  basecopy.getRange('D4').setValue(basecopy.getRange('D4').getValue());
  basecopy.getRange('D5').setValue(basecopy.getRange('D5').getValue());
  basecopy.getRange('D6').setValue(basecopy.getRange('D6').getValue());
  // reset week calculation
  basecopy.getRange('C19').setValue(basecopy.getRange('C19').getValue());
  basecopy.getRange('B20').setValue(basecopy.getRange('B20').getValue());
  basecopy.getRange('B21').setValue(basecopy.getRange('B21').getValue());

  // reset week
  basecopy.getRange('B19').setValue("0");
}

/**
* Returns each weeks starting day and ending day
*
* @return Each weeks starting day and ending day
* @customfunction
*/
function GETDATES() {
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=1 ; i<sheets.length-1 ; i++)
    out.push( [ sheets[i].getName() ] )

  return out;
}

/**
* Returns each weeks total paid amount
*
* @return Each weeks total amount
* @customfunction
*/
function GETWEEKLYAMOUNT() {
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=1 ; i<sheets.length-1 ; i++) out.push( [ sheets[i].getRange('B22').getValue() ] )
  return out;
}

/**
* Returns each weeks paid method or check number
* if paid with a check
*
* @return Each weeks paid method or check number
* @customfunction
*/
function GETCHECKNUMBER() {
  var out = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=1 ; i<sheets.length-1 ; i++) out.push( [ sheets[i].getRange('B23').getValue() ] )
  return out;
}

/**
* Returns time in hours and minutes so pay can be calculated
* based on hours worked.
*
* @param {number} decimaltime Hours in decimal format.
* @return Calculated hours and minutes
* @customfunction
*/
function CONVERTTOCOMPLETETIME(decimaltime) {
  var hrmintime = "";
  if (decimaltime > 0) {
    var hr = Math.floor(decimaltime);
    var min = (decimaltime % hr) * 60;
    hrmintime = hrmintime.concat(hr,":",min.toFixed(2));
  }

  return hrmintime;
}