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

  var insert = ui.createMenu('Insert');
  insert.addItem('Line Above', 'InsertLineAbove');
  insert.addItem('Line Below', 'InsertLineBelow');
  insert.addItem('Buy Side Record', 'InsertBuy');
  insert.addItem('Sell Side Record', 'InsertSell');
  insert.addSeparator();
  insert.addItem('Insert Option', 'ShowInsertOption');
  menu.addSubMenu(insert);

  menu.addSeparator();
  menu.addItem('Convert To Option Symbol', 'Convert2OptionSymbol');
  menu.addSeparator();
  
  menu.addItem('Update Formula', 'UpdateFormula');
  menu.addItem('Update Ticker Summary', 'TickerSummary');
  menu.addItem('Rearrange Data', 'RearrangeData');
  menu.addSeparator();
  menu.addItem('Reformat ML Export', 'ReformatMLExport');

  menu.addToUi();
}

function ShowInsertOption() {
var widget = HtmlService.createHtmlOutputFromFile("optionsDialog.html");
  SpreadsheetApp.getUi().showModalDialog(widget, "Insert Option");
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
  var lastRow=spreadsheet.getRange("O1").getValue();
  var tickers={};
  
  for (var r = 3; r<= lastRow; r++) {
    var val = spreadsheet.getRange("A" + r).getValue();
    tickers[val]=1;
  }
  
  var keys = Object.keys(tickers).sort();
  var r=0;
  for (var key in keys){
    if (keys[key] != "") {
      spreadsheet.getRange("Q" + (2 + r)).setValue(keys[key]);

      // update formulas
      if (r > 0) {
        var copyrange = spreadsheet.getRange('R2:W2');
        var pasterange = spreadsheet.getRange("R" + (2 + r));
        copyrange.copyTo(pasterange);
      }
      r++;
    }
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
  activecell.setNumberFormat("mm/dd/yyyy");
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
  if (activesheet.getRange("B" + activerow).isBlank()) {
    activesheet.getRange("B" + activerow).setValue(GetTodaysDate());
  }
  activesheet.getRange("B" + activerow).setNumberFormat("mm/dd/yyyy");
  // price
  if (activesheet.getRange("C" + activerow).isBlank() && activesheet.getRange("D" + activerow).getValue()!=0) {
    // calculate price from total
    activesheet.getRange("C" + activerow)
    .setValue(activesheet.getRange("F" + activerow).getValue()/
    activesheet.getRange("D" + activerow).getValue());
  } 
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
  activesheet.getRange("H" + activerow).setNumberFormat("mm/dd/yyyy");
  // ticker value lookup
  if (activesheet.getRange("H" + activerow).isBlank()) {
    activesheet.getRange("H" + activerow).setValue('=GOOGLEFINANCE(A' + activerow + ')');
  }
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
  // A - above
  InsertLine('A');
}

function InsertLineBelow() {
  // B - below
  InsertLine('B');
}

function InsertBuySellRecord(side) {
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //var activeRange = activesheet.getActiveRange();
  var row = activesheet.getActiveCell().getRow();
  var startRow = 1;
  var endRow = 1;
  var sideColStart = "B";// buy side
  
  // sell side
  if (side == 'S') {
    sideColStart = "G"; // sell side
  }
  
  if (activesheet.getRange("A"+ row).isBlank()) {
    Browser.msgBox("Please select the correct line item.");
    return;
  }

  // check has other records (if date is blank)
  if (activesheet.getRange( sideColStart + row).isBlank()) {
    InsertLineBelow();
    var lastRow = activesheet.getLastRow();
    while ((row + endRow) < lastRow){
      if (activesheet.getRange(sideColStart + (row + endRow + 1)).isBlank()){
        break;
      }
      else {
        endRow++;
      }
    }
  }
  else {
    // only a single record
    // insert 2 new lines; existing and new record
    InsertLineBelow();
    InsertLineBelow();

    //copy record
    endRow = endRow + 1
    // date
    activesheet.getRange(sideColStart + (row + endRow)).setValue(activesheet.getRange(sideColStart + row).getValue());
    // price
    activesheet.getRange("C"+ (row + endRow)).setValue(activesheet.getRange("C"+ row).getValue());
    // quantity
    activesheet.getRange("D"+ (row + endRow)).setValue(activesheet.getRange("D"+ row).getValue());
    // fees
    activesheet.getRange("E"+ (row + endRow)).setValue(activesheet.getRange("E"+ row).getValue());
    activesheet.getRange("E"+ (row + endRow)).setNumberFormat(activesheet.getRange("E"+ row).getNumberFormat());
    
    // total
    var copyrange = activesheet.getRange('F'+ row);
    var pasterange = activesheet.getRange('F' + (row + endRow));
    copyrange.copyTo(pasterange);

    // remove date
    activesheet.getRange(sideColStart + row).clearContent();
  }

  // add/update formulas
  activesheet.getRange("F"+ row).setFormula("=sum(F" + (row + startRow) + ":F" + (row + endRow) + ")");
  activesheet.getRange("E"+ row).setFormula("=sum(E" + (row + startRow) + ":E" + (row + endRow) + ")");
  activesheet.getRange("D"+ row).setFormula("=sum(D" + (row + startRow) + ":D" + (row + endRow) + ")");
  activesheet.getRange("C"+ row).setFormula("F" + row + "/D" + row);
}

function InsertBuy() {
  InsertBuySellRecord('B');
}

function InsertSell() {
  InsertBuySellRecord('S');
}

function InsertLine(dir) {
  // Direction
  // A - Above
  // B - Below
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

function DeleteEmptyRows() {
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = 2;

  do {
    var selectRange = activesheet.getRange("A" + row + ":P" + row);
     if (activesheet.getRange("D" + row).isBlank()) {
       selectRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
     }
  } while(activesheet.getRange("D" + row + 1).isBlank())
}

function CleanCancelledRecords(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = 2;

  do {
    var selectRange = activesheet.getRange("A" + row + ":P" + row);
    if (activesheet.getRange("D" + row).getValue() == "Cancelled") {
      selectRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
    else
      row++;
  } while (!activesheet.getRange("D" + row + 1).isBlank())

}

function Convert2OptionSymbol() {
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeCell = activesheet.getSelection().getCurrentCell();
  var activeRange = activesheet.getActiveRange();
  var rows = activeRange.getNumRows();
  var row = 0;
  do {
    //var ndx = activeCell.offset(row,0).getValue().indexOf('\n');
    var option = activeCell.offset(row,0).getValue().replace('\t',' ').replace('\n', ' ').trim();
    //option = option.slice(0,ndx) + " " + option.slice(ndx);

    // (\d{4}.\d*.\d{2})|(\d{2}.\d*.\d{4})|(\d{2}[^A-Za-z0-9.]\d{2}|[A-z]{3}\s\d{2},.\d{4})
    var arrdate = option.match(/(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})|(\s[A-Z]{3}.\d{1,2},.\d{4})|(\s\d{1,2}.[A-Z]{3}.\d{1,2})/gi);
    var date = new Date(arrdate[0].trim());
    date.toLocaleDateString("en-US",{month: '2-digit', day: '2-digit', year: 'numeric'})
    var sdate = date.getFullYear().toString().substring(2) + (date.getMonth()+1).toString().padStart(2,'0') + date.getDate().toString().padStart(2,'0')
    var ldate = [(date.getMonth()+1).toString().padStart(2,'0'), date.getDate().toString().padStart(2,'0'), date.getFullYear()].join('/');
    var ticker = option.match(/([A-Z]{1,}\s)/);
    var price = option.match(/(\$\d*.\d{2})|(\s\d*.00)/);
    var ndx = price[0].indexOf('$');
    if (ndx == -1) {
      price[0] = '$' + price[0].trim();
    }
    var ndx = price[0].indexOf('.');
    if (ndx == -1) {
      price[0] = price[0].trim() + ".00";
    }
    var type = option.match(/(Call|Put)/i);
    type[0] = type[0].charAt(0).toUpperCase() + type[0].slice(1).toLowerCase();
    
    // set ticker
    activeCell.offset(row,-1).setValue(ticker[0].trim());
    //Logger.log(ticker[0].trim());
    
    // reformat option
    activeCell.offset(row,0).setValue([ticker[0].trim(),ldate,type[0],price[0]].join(' '));
    //Logger.log([ticker[0].trim(),ldate,type[0],price[0]].join(' '));
    
    // set option code
    activeCell.offset(row,1).setValue(
      [ticker[0].trim(),sdate,type[0].substring(0,1),price[0].replace('$','').replace('.','').padStart(7,'0'),"0"].join(''));
    //Logger.log([ticker[0].trim(),sdate,type[0].substring(0,1),price[0].replace('$','').replace('.','').padStart(7,'0'),"0"].join(''));
    
    // set type
    activeCell.offset(row,2).setValue(type[0]);
    //Logger.log(type[0]);
    
    // set option code
    //activeCell.offset(row,1).setValue(values[0] + date.getFullYear().toString().substring(2) + (date.getMonth()+1).toString().padStart(2,'0') + date.getDate().toString().padStart(2,'0') + values[2 + vOffset].substring(0,1) + values[3 + vOffset].toString().replace('$','').replace('.','').padStart(7,'0') + "0");

      row++;
  } while (row < rows)
}

function UpdateFormula() {
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRange = activesheet.getActiveRange();
  var rows = activeRange.getNumRows();
  var row = activesheet.getActiveCell().getRow();
  var activeCell = activesheet.getSelection().getCurrentCell();
  var offset = 0;
  
  do {
    var val = activeCell.offset(offset,0).getFormula();
    // no formula
    if (val =="") {
      val = "=(" + activeCell.offset(offset,0).getValue() + ")/I" + (row + offset);
      activeCell.offset(offset,0).setFormula(val);
      activesheet.getRange(activeCell.getRow() + offset , activeCell.getColumn()).setNumberFormat("$0.00");
    }
    offset++;
  } while (offset < rows )
}

function CalculateOptionFees(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = activesheet.getActiveCell().getRow();
  var col = activesheet.getActiveCell().getColumn();
  var buycol = "K".charCodeAt(0) - "A".charCodeAt(0) + 1
  var sellcol = "O".charCodeAt(0) - "A".charCodeAt(0) + 1
  var numContractsCol = "I".charCodeAt(0) - "A".charCodeAt(0) + 1
  var tradeAmountCol = "J".charCodeAt(0) - "A".charCodeAt(0) + 1

  var acc = activesheet.getRange(row, 1).getValue();
  var numContracts = activesheet.getRange(row, numContractsCol).getValue();
  var tradeAmount = activesheet.getRange(row, tradeAmountCol).getValue() * 100;
  var tradeCount = (activesheet.getActiveCell().offset(0,-1).getFormula().match(/\+/g) || []).length + 1;

  var regulatoryTransactionFee = 0;
  var tradingActivityFee = 0;
  var optionsRegulatoryFee = 0;
  var clearingFee = 0;
  var numSales = 0;

  if (acc === "WB-O" || acc === "WB-Y") {
    // https://www.webull.com/pricing

    if (col == sellcol) {
      regulatoryTransactionFee = 0.0000229 * tradeAmount * tradeCount;
      regulatoryTransactionFee = Math.max(regulatoryTransactionFee,0.01); // Min $0.01
      tradingActivityFee = 0.00218 * numContracts * tradeCount;
      tradingActivityFee = Math.max(tradingActivityFee, 0.01); //Min $0.01
      var sellfees = regulatoryTransactionFee + tradingActivityFee;
    }
    
    if (col == sellcol || col == buycol) {
      optionsRegulatoryFee = 0.01815 * numContracts;
      clearingFee = 0.02 * numContracts;
      clearingFee = Math.min(55, clearingFee); // (Max $55 per Trade)
      var fees = optionsRegulatoryFee + clearingFee;
    }
  }
  else if ( acc === "RH") {
    if (col == sellcol) {
      regulatoryTransactionFee = 0.00244 * numContracts;
      regulatoryTransactionFee = Math.min(7.27, regulatoryTransactionFee); // (Max $7.27 per Trade)
    }
  }
  else if ( acc === "ML") {
    if (col == sellcol) {
      regulatoryTransactionFee = 0.01;
    }
    if (col == sellcol || col == buycol) {
      optionsRegulatoryFee = 0.65 * numContracts;
    }
  }
  else if ( acc === "TS") {
    if (col == buycol) {
      regulatoryTransactionFee = 0.01;
    }
    if (col == sellcol) {
      regulatoryTransactionFee = 0.02;
    }
    if (col == sellcol || col == buycol) {
      optionsRegulatoryFee = 0.65 * numContracts;
    }
  }
  return (regulatoryTransactionFee + tradingActivityFee + optionsRegulatoryFee + clearingFee);
}

function InsertOption(form){
  var row = [form.name, form.feedback, form.rating];
}

function ClearMLReformatedData(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = 2;

  do {
    var opendate = activesheet.getRange("N" + row).getValue();
    var closedate = activesheet.getRange("O" + row).getValue();
    activesheet.getRange("N" + row + ":Z" + row).clear();
    row = row + 1;
  } while (opendate !== "" && closedate !== "")

}

function ReformatMLExport(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = activesheet.getActiveCell().getRow();
  var col = activesheet.getActiveCell().getColumn();

  ClearMLReformatedData();

  if (col == 2) {
    var orders = new Map();
    do {
      var date = activesheet.getRange("B" + row).getValue().match(/\d{1,2}\/\d{1,2}\/\d{2,4}/);
      var action = activesheet.getRange("E" + row).getValue();
      var arrAction = action.match(/(Sell|Buy)/i);
      action = arrAction[0];
      var optioncall = activesheet.getRange("F" + row).getValue();
      var qty = activesheet.getRange("G" + row).getValue();
      var price = activesheet.getRange("L" + row).getValue();
      var arrPrice = price.match(/\d{1,3}.\d{2}/);

      if (Array.isArray(arrPrice)) {
        price = arrPrice[0];
        var princestring = price;
        if (qty > 1) {
          princestring = price + " * " + qty;
        }
        if (orders.has(optioncall)){

          if (action == "Buy") {
            orders.get(optioncall).BuyQty = orders.get(optioncall).BuyQty + " + " + qty;
            orders.get(optioncall).BuyPrice = orders.get(optioncall).BuyPrice + " + " + princestring;
          }
          if (action == "Sell") {
            orders.get(optioncall).SellDate = date[0];
            orders.get(optioncall).SellQty = orders.get(optioncall).SellQty + " + " + qty;
            orders.get(optioncall).SellPrice = orders.get(optioncall).SellPrice + " + " + princestring;
          }
        }
        else {
          if (action == "Buy") {
            orders.set(optioncall,{Type : "Buy", Action : optioncall.match(/(CALL|PUT)/i)[0], BuyDate : date[0], BuyQty : qty, BuyPrice : princestring, SellDate: "", SellQty : "0", SellPrice : "0"});
          }
          if (action == "Sell") {
            orders.set(optioncall,{Type : "Sell", Action : optioncall.match(/(CALL|PUT)/i)[0], BuyQty : "0", BuyPrice : "0", SellDate : date[0], SellQty : qty, SellPrice : princestring});
          }
        }
      }

      row = row - 1;
    } while (row > 5)

    row = 2;
    for ( let [id,order] of orders.entries()){
      activesheet.getRange("N" + row).setValue(order.BuyDate);
      activesheet.getRange("O" + row).setValue(order.SellDate);
      activesheet.getRange("Q" + row).setValue(id);
      activesheet.getRange("S" + row).setValue(order.Action);
      activesheet.getRange("T" + row).setValue(order.Type);
      var qty = order.BuyQty;
      if (qty.length > 4) {
        qty = qty.replace(/^0\s\+\s/,"")
      }
      activesheet.getRange("U" + row).setValue("=" + qty);
      var price = order.BuyPrice.replace(/^0\s\+\s/,"");
      activesheet.getRange("V" + row).setValue(price);
      qty = order.SellQty;
      if (qty.length > 4) {
        qty = qty.replace(/^0\s\+\s/,"")
      }
      activesheet.getRange("Y" + row).setValue("=" + qty);
      price = order.SellPrice.replace(/^0\s\+\s/,"");
      activesheet.getRange("Z" + row).setValue(price);
      row = row + 1;
    }
  }
}

function GetNextFridayDate(){
  var date = new Date("1/7/2023");
  var day = date.getDay();
  // Logger.log(day + " " + date.toString());
  //day = (day == 6) : 0, day;
  var offset = 5 - date.getDay();
  date.setDate(date.getDate() + offset);
  day = date.getDay();
  
  Logger.log(day + " " + date.toDateString());
  // return date.toDateString();
}
