function onOpen() {   
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Investments');

  menu.addItem('Hyperlink Me', 'HyperlinkMe');
  menu.addItem('Fill Current Line', 'FillCurrentLine');
  menu.addItem('Fill Current Line Buy Side', 'FillCurrentLineBuySide');
  menu.addItem('Fill Current Line Sell Side', 'FillCurrentLineSellSide');
  menu.addItem('Update Ticker Summary', 'TickerSummary');
  menu.addItem('Rearrange Data', 'RearrangeData');
  menu.addToUi();
}

function CopyTodaysRates() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.getRange('c10').setValue(GetTodaysDate());
  spreadsheet.getRange('c11').setValue(spreadsheet.getRange('b11').getValue());
  spreadsheet.getRange('c12').setValue(spreadsheet.getRange('b12').getValue());
  spreadsheet.getRange('c13').setValue(spreadsheet.getRange('b13').getValue());
}

function RearrangeData(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var iStartRow = 3
  var iLastRow = GetLastRow(spreadsheet, iStartRow);
  
  for (var i=iStartRow; i<= iLastRow; i++) {
    var ticker = spreadsheet.getRange('A' + i).getValue();
  }
}

function GetLastRow(sheet, startRow){
  var r = startRow
  do {
    r++;
  } while (sheet.getRange('C' + r).getValue() != '');
  
  return (r-1);
}

function HyperlinkMe(){
  var val = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue();
  var url= "=HYPERLINK(\"https://www.marketwatch.com/investing/stock/" + val.toLowerCase() + "\",\"" + val.toUpperCase() + "\")";
  SpreadsheetApp.getActiveSpreadsheet().getActiveCell().setValue(url);
}

function TickerSummary(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastRow=spreadsheet.getRange("P2").getValue();
  var tickers={};
  
  for (var r = 3; r<= lastRow; r++) {
    var val = spreadsheet.getRange("A" + r).getValue();
    tickers[val]=1;
  }
  
  var keys = Object.keys(tickers);
  var r=0;
  for (var key in keys){
    spreadsheet.getRange("Q" + (5 + r)).setValue(keys[key]);
    r++;
  }
}

function GetTodaysDate(){
  var today = new Date();
  var dd = today.getDate();
  var mm = today.getMonth() + 1; //January is 0!
  var yyyy = today.getFullYear();
  
  today = mm + '/' + dd + '/' + yyyy;
  
  return today;
}

function FillTodaysDate(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet();
  var activecell = activesheet.getActiveCell();

  activecell.setValue(GetTodaysDate());
}

function FillCurrentLine(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet();
  var activecell = activesheet.getActiveCell();
  var activerow = activecell.getRow();
  
  HyperlinkMe();
  FillCurrentLineBuySide();
  FillCurrentLineSellSide();
  
  // gain/loss
  var val='=K' + activerow + '-F' + activerow;
  activesheet.getRange("L" + activerow).setValue(val);
  activesheet.getRange("L" + activerow).setNumberFormat("$0.00");
  // gain/loss percentage
  var copyrange = activesheet.getRange('M3');
  var pasterange = activesheet.getRange('M' + activerow);
  copyrange.copyTo(pasterange);
  
  // status
  var val='=if(G' + activerow + '="","Open","")';
  activesheet.getRange("N" + activerow).setValue(val);
  activesheet.getRange("N" + activerow).setHorizontalAlignment('center');
}

function FillCurrentLineBuySide(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet();
  var activecell = activesheet.getActiveCell();
  var activerow = activecell.getRow();
  
  // Buy side
  var today = new Date();
  activesheet.getRange("B" + activerow).setValue(GetTodaysDate());
  // price
  activesheet.getRange("C" + activerow).setNumberFormat("$0.00");
  // fees
  activesheet.getRange("E" + activerow).setValue('0');
  activesheet.getRange("E" + activerow).setNumberFormat("$0.00");
  // total
  var val='=C' + activerow + '*D' + activerow + '+E' + activerow;
  activesheet.getRange("F" + activerow).setValue(val);
  activesheet.getRange("F" + activerow).setNumberFormat("$0.00");
}

function FillCurrentLineSellSide(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet();
  var activecell = activesheet.getActiveCell();
  var activerow = activecell.getRow();

  // sell side
  // ticker value lookup
  activesheet.getRange("H" + activerow).setValue('=GOOGLEFINANCE(A' + activerow + ')');
  activesheet.getRange("H" + activerow).setNumberFormat("$0.00");
  // qty
  activesheet.getRange("I" + activerow).setValue('=D'+ activerow);
    // fees
  activesheet.getRange("J" + activerow).setValue('0');
  activesheet.getRange("J" + activerow).setNumberFormat("$0.00");
    // total
  var val='=H' + activerow + '*I' + activerow + '-J' + activerow;
  activesheet.getRange("K" + activerow).setValue(val);
  activesheet.getRange("K" + activerow).setNumberFormat("$0.00");
}