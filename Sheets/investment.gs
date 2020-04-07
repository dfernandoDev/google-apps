function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Investments');

  var hyper = ui.createMenu('Hyperlink');
  hyper.addItem('Yahoo', 'HyperlinkToYahoo');
  hyper.addItem('Market Watch', 'HyperlinkToMarketWatch');
  menu.addSubMenu(hyper);

  var fill = ui.createMenu('Fill');
  fill.addItem('Today\'s Date', 'FillTodaysDate');
  fill.addItem('Current Line', 'FillCurrentLine');
  fill.addItem('Current Line Buy Side', 'FillCurrentLineBuySide');
  fill.addItem('Current Line Sell Side', 'FillCurrentLineSellSide');
  menu.addSubMenu(fill);

  menu.addItem('Update Ticker Summary', 'TickerSummary');
  menu.addItem('Rearrange Data', 'RearrangeData');

  var insert = ui.createMenu('Insert');
  insert.addItem('Line Above', 'InsertLineAbove');
  insert.addItem('Line Below', 'InsertLineBelow');
  menu.addSubMenu(insert);

  menu.addToUi();
}

function CopyTodaysRates() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  spreadsheet.getRange('c10').setValue(GetTodaysDate());
  spreadsheet.getRange('c11').setValue(spreadsheet.getRange('b11').getValue());
  spreadsheet.getRange('c12').setValue(spreadsheet.getRange('b12').getValue());
  spreadsheet.getRange('c13').setValue(spreadsheet.getRange('b13').getValue());
}

function RearrangeData(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 3
  //var lastRow = spreadsheet.getRange('P2').getValue();
  var lastRow = GetLastRow(spreadsheet, startRow);

  var currentRow = lastRow;
  var lastCloseRow = lastRow + 1;

  do {
    var ticker = spreadsheet.getRange('A' + currentRow).getValue();
    var sdate = spreadsheet.getRange('G' + currentRow).getValue();
    var openclose = spreadsheet.getRange('N' + currentRow).getValue();
    // find open stocks
    if (ticker.length > 0 && openclose.length > 0 && currentRow != lastRow) {
      var endOpenRow = currentRow;
      do {
        endOpenRow++;
        // check if this is a single line data
        ticker = spreadsheet.getRange('A' + endOpenRow).getValue();
        var bQty = spreadsheet.getRange('D' + endOpenRow).getValue();
        var sQty = spreadsheet.getRange('I' + endOpenRow).getValue();
        var temp1a = ticker.length;
        var temp1 = ticker.length > 0;
        var temp2 = bQty > 0;
        var temp3 = sQty > 0;
        var temp4 = (!(ticker.length > 0 && bQty > 0 && sQty > 0));

      } while (!(ticker.length > 0 && bQty > 0 && sQty > 0));

      // copy line(s) to below
      endOpenRow--;
      var sourceRange = spreadsheet.getRange("A" + currentRow + ":O" + endOpenRow);
      var destinationRange = spreadsheet.getRange("A" + lastCloseRow);
      sourceRange.copyTo(spreadsheet.getRange("A" + lastCloseRow));

      sourceRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
    currentRow--;
  } while (currentRow >= startRow);
}

function GetLastRow(sheet, startRow){
  var r = startRow
  do {
    r++;
    var ticker = sheet.getRange('A' + r).getValue();
    var qBought = sheet.getRange('D' + r).getValue();
    var qSold = sheet.getRange('I' + r).getValue();
  } while (ticker != '' || qBought != '' || qSold != '');

  return (r-1);
}

function HyperlinkToYahoo(){
  HyperlinkTo("yahoo");
}

function HyperlinkToMarketWatch(){
  HyperlinkTo("marketwatch");
}

function HyperlinkTo(site){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = spreadsheet.getActiveCell().getRow();
  var val = spreadsheet.getRange('A' + row).getValue();
  // var val = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue();
  var url ="https://";

  switch (site) {
    case "yahoo":
      url =  url.concat("finance.yahoo.com/quote/");
      break;
    case "marketwatch":
      url =  url.concat("www.marketwatch.com/investing/stock/");
      break;
  }

  // var url = "https://finance.yahoo.com/quote/" + val.toLowerCase();
  // var url ="https://www.marketwatch.com/investing/stock/" + val.toLowerCase();

  url = url.concat(val.toLowerCase());
  var hyper= "=HYPERLINK(\"" + url + "\",\"" + val.toUpperCase() + "\")";

  // SpreadsheetApp.getActiveSpreadsheet().getActiveCell().setValue(hyper);
  spreadsheet.getRange('A' + row).setValue(hyper);
}

function TickerSummary(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activecell = activesheet.getActiveCell();

  activecell.setValue(GetTodaysDate());
}

function FillCurrentLine(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activecell = activesheet.getActiveCell();
  var activerow = activecell.getRow();

  HyperlinkToYahoo();
  // HyperlinkToMarketWatch();
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
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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

function InsertLineAbove() {
  InsertLine('A');
}

function InsertLineBelow() {
  InsertLine('B');
}

function InsertLine(dir) {
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRange = activesheet.getActiveRange();
  var row = activesheet.getActiveCell().getRow();

  switch (dir) {
    case 'A':
      break;
    case 'B':
      row += 1;
      break;
  }

  var selectRange = activesheet.getRange("A" + row + ":O" + row);
  selectRange.activate();

  selectRange.insertCells(SpreadsheetApp.Dimension.ROWS);
  activesheet.setActiveRange(activeRange)

}
