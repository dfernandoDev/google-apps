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
  
  insert = ui.createMenu('ML');
  insert.addItem('Clear ML Export', 'ClearMLExport');
  insert.addItem('Reformat ML Order Export', 'ReformatMLOrderExport');
  insert.addItem('Reformat ML Activity Export', 'ReformatMLActivityExport');
  menu.addSubMenu(insert);
  
  insert = ui.createMenu('RH');
  insert.addItem('Clear RH Export', 'ClearRHExport');
  insert.addItem('Reformat RH Export', 'ReformatRHExport');
  menu.addSubMenu(insert);

  insert = ui.createMenu('WB');
  insert.addItem('Clear WB Export', 'ClearWBExport');
  insert.addItem('Reformat WB Export', 'ReformatWBExport');
  menu.addSubMenu(insert);
  
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
    var option = activeCell.offset(row,0).getValue().replace('\t',' ').replace('\n', ' ').replace("00:00:00 EST","").replace("00:00:00 EDT","").trim();
    //option = option.slice(0,ndx) + " " + option.slice(ndx);

    // (\d{4}.\d*.\d{2})|(\d{2}.\d*.\d{4})|(\d{2}[^A-Za-z0-9.]\d{2}|[A-z]{3}\s\d{2},.\d{4})
    var arrdate = option.match(/(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})|(\s[A-Z]{3}.\d{1,2},.\d{4})|(\s\d{1,2}.[A-Z]{3}.\d{1,2})/gi);
    var date = new Date(arrdate[0].trim());
    date.toLocaleDateString("en-US",{month: '2-digit', day: '2-digit', year: 'numeric'})
    var sdate = date.getFullYear().toString().substring(2) + (date.getMonth()+1).toString().padStart(2,'0') + date.getDate().toString().padStart(2,'0')
    var ldate = [(date.getMonth()+1).toString().padStart(2,'0'), date.getDate().toString().padStart(2,'0'), date.getFullYear()].join('/');
    var ticker = option.match(/([A-Z]{1,}\s)/);
    var price = option.match(/(\$\d*.\d{2})|(\s\d[^\S]*.\d*.[05]0)/);
    price[0]=price[0].replace(',','');
    var ndx = price[0].indexOf('$');
    if (ndx == -1) {
      price[0] = '$' + price[0].trim();
    }
    else {
      price[0] = price[0].replace(' ','').trim();
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
    let formula = activeCell.offset(offset,0).getFormula();
    let val = activeCell.offset(offset,0).getValue();
    let plus = val.toString().indexOf('+');
    let multi = val.toString().indexOf('*');
    // no formula
    if (formula =="" && (plus >= 0 || multi >=0)) {
      formula = "=(" + activeCell.offset(offset,0).getValue() + ")/I" + (row + offset);
      activeCell.offset(offset,0).setFormula(formula);
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
  var tradeSymbolCol = "D".charCodeAt(0) - "A".charCodeAt(0) + 1

  var acc = activesheet.getRange(row, 1).getValue();
  var tradeSymbol = activesheet.getRange(row, tradeSymbolCol).getValue();
  var numContracts = activesheet.getRange(row, numContractsCol).getValue();
  var tradeAmount = activesheet.getRange(row, tradeAmountCol).getValue() * 100;
  var tradeCount = (activesheet.getActiveCell().offset(0,-1).getFormula().match(/\+/g) || []).length + 1;

  var regulatoryTransactionFee = 0;
  var tradingActivityFee = 0;
  var optionsRegulatoryFee = 0;
  var clearingFee = 0;
  var numSales = 0;
  // webull only
  var exchangeProprietaryFee = 0;
  var contractFee = 0;

  if (acc === "WB-O" || acc === "WB-Y") {
    // https://www.webull.com/pricing

    if (tradeSymbol == "SPX"){
      contractFee = 0.55 * numContracts;
      exchangeProprietaryFee = 0.66 * numContracts;
      //return (contractFee + exchangeProprietaryFee);
    }
    else if (tradeSymbol == "SPXW"){
      contractFee = 0.55 * numContracts;
      exchangeProprietaryFee = 0.58 * numContracts;
    }

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
  //Logger.log (regulatoryTransactionFee + tradingActivityFee + optionsRegulatoryFee + clearingFee + exchangeProprietaryFee + contractFee);
  return (regulatoryTransactionFee + tradingActivityFee + optionsRegulatoryFee + clearingFee + exchangeProprietaryFee + contractFee);
}

function InsertOption(form){
  var row = [form.name, form.feedback, form.rating];
}

function ClearReformatedData(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = 2;

  do {
    var opendate = activesheet.getRange("N" + row).getValue();
    var closedate = activesheet.getRange("O" + row).getValue();
    row = row + 1;
  } while (opendate !== "" || closedate !== "")

  activesheet.getRange("M2:Z" + row).clear();
}

function GetOrderLastRow(row, col){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  do {
    row = row + 1;
  } while (activesheet.getRange(row, col).getValue() !== "")

  return row - 1;
}

function ReplaceBeginingZero(val){
  let hasZero = val.toString().match(/^0\s\+\s/,"");
  let ret = val;
  if (hasZero !== null) {
    ret = val.toString().replace(/^0\s\+\s/,"");
  }

  return ret;
}

function AddNewOrderItem(orders, optioncall, optionsymbol, date, qty, price, action) {
  if (action == "Buy") {
    orders.set(optioncall,
      {
        Symbol : optionsymbol, 
        Type : "Buy", 
        Action : optioncall.match(/(CALL|PUT)/i)[0], 
        BuyDate : date, 
        BuyQty : qty, 
        BuyQtyInt : qty, 
        BuyPrice : price, 
        SellDate: "", 
        SellQty : "0", 
        SellQtyInt : 0, 
        SellPrice : "0"
      }
    );
  }
  else if (action == "Sell") {
    orders.set(optioncall,
      {
        Symbol : optionsymbol, 
        Type : "Sell", 
        Action : optioncall.match(/(CALL|PUT)/i)[0], 
        BuyQty : "0", 
        BuyQtyInt : 0, 
        BuyPrice : "0", 
        SellDate : date, 
        SellQty : qty, 
        SellQtyInt : qty, 
        SellPrice : price
      }
    );
  }
}

function AddOrderItem(orders, orderIDs, optioncall, optionsymbol, date, qty, price, action) {
  let key = optioncall;
  if (orderIDs.get(optioncall).Count > 1){
    key = optioncall + "#" + orderIDs.get(optioncall).Count;
  }
  if (orders.get(key).BuyQtyInt > 0 && orders.get(key).BuyQtyInt == orders.get(key).SellQtyInt) {
    if (orderIDs.has(optioncall)) {
      orderIDs.set(optioncall,{ Count : orderIDs.get(optioncall).Count + 1});
      key = optioncall + "#" + orderIDs.get(optioncall).Count;
      AddNewOrderItem(orders, key, optionsymbol, date, qty, price, action);
    }
  }
  else if (action == "Buy") {
    orders.get(key).BuyQty = orders.get(key).BuyQty + " + " + qty;
    orders.get(key).BuyQtyInt = orders.get(key).BuyQtyInt + qty;
    orders.get(key).BuyPrice = orders.get(key).BuyPrice + " + " + price;
  }
  else if (action == "Sell") {
    orders.get(key).SellDate = date;
    orders.get(key).SellQty = orders.get(key).SellQty + " + " + qty;
    orders.get(key).SellQtyInt = orders.get(key).SellQtyInt + qty;
    orders.get(key).SellPrice = orders.get(key).SellPrice + " + " + price;
  }  
}

function BuildOrderMap(orders, orderIDs, optioncall, optionsymbol, date, qty, price, action){
  if (orders.has(optioncall)){
    AddOrderItem(orders, orderIDs, optioncall, optionsymbol, date, qty, price, action);
  }
  else {
    orderIDs.set(optioncall,{ Count : 1});
    AddNewOrderItem(orders, optioncall, optionsymbol, date, qty, price, action);
  }
}

function PopulateFormatedData(orders, account) {
  let activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let row = 2;
  for ( let [id,order] of orders.entries()){
    activesheet.getRange("M" + row).setValue(account)
    activesheet.getRange("N" + row).setValue(order.BuyDate);
    activesheet.getRange("O" + row).setValue(order.SellDate);
    activesheet.getRange("Q" + row).setValue(id.split("#")[0]);
    activesheet.getRange("R" + row).setValue(order.Symbol);
    activesheet.getRange("S" + row).setValue(order.Action);
    activesheet.getRange("T" + row).setValue(order.Type);
    let qty = ReplaceBeginingZero(order.BuyQty);
    activesheet.getRange("U" + row).setValue("=" + qty);
    let price = ReplaceBeginingZero(order.BuyPrice);
    activesheet.getRange("V" + row).setValue(price);
    qty = ReplaceBeginingZero(order.SellQty);
    activesheet.getRange("Y" + row).setValue("=" + qty);
    price = ReplaceBeginingZero(order.SellPrice);
    activesheet.getRange("Z" + row).setValue(price);
    row = row + 1;
  }
}

function ReformatMLOrderExport(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var col = 2; //activesheet.getActiveCell().getColumn();
  //var row = activesheet.getActiveCell().getRow();
  var row = GetOrderLastRow(5 ,col);

  var orders = new Map();
  var orderIDs = new Map();
  do {
    var date = activesheet.getRange("B" + row).getValue().match(/\d{1,2}\/\d{1,2}\/\d{2,4}/);
    var action = activesheet.getRange("E" + row).getValue();
    var arrAction = action.match(/(Sell|Buy)/i);
    action = arrAction[0];
    var optioncall = activesheet.getRange("F" + row).getValue();
    var qty = activesheet.getRange("G" + row).getValue();
    var price = activesheet.getRange("L" + row).getValue();
    var arrPrice = price.match(/\d{1,3}.\d{1,2}/);

    if (Array.isArray(arrPrice)) {
      price = arrPrice[0];
      var princestring = price;
      if (qty > 1) {
        princestring = price + " * " + qty;
      }
      
      BuildOrderMap(orders, orderIDs, optioncall, "", date, qty, price, action);
    }

    row = row - 1;
  } while (row > 5)

  ClearReformatedData();
  PopulateFormatedData(orders, "ML");
}

function ReformatMLActivityExport(){
  let activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let col = 1;
  let firstDataRow = 9;
  let row = GetOrderLastRow(firstDataRow ,col);

  let orders = new Map();
  var orderIDs = new Map();
  do {
    let price = activesheet.getRange("H" + row).getValue();
    let date = activesheet.getRange("A" + row).getValue();//.match(/\d{1,2}\/\d{1,2}\/\d{2,4}/);
    let action = activesheet.getRange("D" + row).getValue();
    let arrAction = action.match(/(Sale|Purchase)/i);
    if (Array.isArray(arrAction)) {
      action = arrAction[0];
    }
    else {
      arrAction = action.match(/(Expired)/i);
      if (Array.isArray(arrAction)) {
        arrAction = "Sale"
        price = 0;
      }
    }
    action = action.replace("Purchase","Buy");
    action = action.replace("Sale","Sell");
    let optioncall = activesheet.getRange("D" + row).getValue();
    let arrOptionCall = optioncall.match(/(Call|Put).\w+.\d{5}|Exp.\d{2}-\d{2}-\d{2}/gi);
    let arrTypeSymbol = arrOptionCall[0].split(' ')
    optioncall = arrTypeSymbol[1] + arrOptionCall[1].replace("EXP ", " ").replaceAll("-","/") + " " + arrTypeSymbol[0] + " $" + (arrTypeSymbol[2]*100/100);
    let qty = activesheet.getRange("G" + row).getValue();
    
    BuildOrderMap(orders, orderIDs, optioncall, "", date, qty, price, action);
    row = row - 1;
  } while (row > firstDataRow-1)
  ClearReformatedData();
  PopulateFormatedData(orders, "ML");
}

function ReformatRHExport(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var col = 1; //activesheet.getActiveCell().getColumn();
  //var row = activesheet.getActiveCell().getRow();
  var row = 3;

  do {
    row = row + 1;
  } while (activesheet.getRange(row, col).getValue() !== "" || activesheet.getRange(row + 1, col).getValue() !== "")

  row = row - 1;

  var orders = new Map();
  var orderIDs = new Map();
  do {
    var contract = activesheet.getRange(row, col).getValue();
    if (contract == "Canceled") {
      row = row - 4;
    }
    else {
      var date = activesheet.getRange(row, col).offset(-2,0).getValue();
      if (date.toString().indexOf(' ') == -1) {
        date = new Date();
      }
      var option = activesheet.getRange(row, col).offset(-3,0).getValue();
      var action = option.match(/(Sell|Buy)/i)[0];
      var qty = contract.match(/\d{1,2} /)[0].trim(); 
      var price = contract.match(/[$]\d{1,2}.\d{2}/g)[0].replace('$','');
      var arrStrikePrice = option.match(/[$]\d{1,4}[.]\d|[$]\d{1,4}/);
      //let strikePrice = arrStrikePrice[0].replace('$','');

      var optioncall = option.match(/\s\w+/)[0].trim() + ' ' + option.match(/\d{1,2}\/\d{1,2}/)[0] + '/2023 ' + option.match(/Call|Put/i)[0] + ' $' + Number(arrStrikePrice[0].replace('$','')).toFixed(2);

      var princestring = price;
      if (qty > 1) {
        princestring = price + " * " + qty;
      }

      BuildOrderMap(orders, orderIDs, optioncall, "", date, qty, price, action);

      row = row - 5;
    }
  } while (row > 3)

  ClearReformatedData();
  PopulateFormatedData(orders, "RH");
}

function ReformatWBExport(){
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var col = 1; //activesheet.getActiveCell().getColumn();
  //var row = activesheet.getActiveCell().getRow();
  var row = GetOrderLastRow(5 ,col);

  var orders = new Map();
  var orderIDs = new Map();

  do {
    var optioncall = activesheet.getRange("A" + row).getValue();
    var optionsymbol = activesheet.getRange("B" + row).getValue();
    var action = activesheet.getRange("C" + row).getValue();
    var arrAction = action.match(/(Sell|Buy)/i);
    action = arrAction[0];
    var status = activesheet.getRange("D" + row).getValue();
    var qty = activesheet.getRange("E" + row).getValue();
    var price = activesheet.getRange("H" + row).getValue();
    //var arrPrice = price.match(/\d{1,3}.\d{2}/);
    var date = activesheet.getRange("J" + row).getValue().match(/\d{1,2}\/\d{1,2}\/\d{2,4}/);

    if (status == "Filled") {
      var princestring = price;
      if (qty > 1) {
        princestring = price + " * " + qty;
      }

      BuildOrderMap(orders, orderIDs, optioncall, optionsymbol, date[0], qty, princestring, action);
    }
    row = row - 1;
  } while (row > 1)

  ClearReformatedData();
  PopulateFormatedData(orders, "WB");
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

function ClearWBExport(){
  ClearImportedData(1,"A")
  ClearReformatedData();
}

function ClearRHExport(){
  ClearImportedData(1,"A")
  ClearReformatedData();
}

function ClearMLExport(){
  ClearImportedData(5,"B")
  ClearReformatedData();
}

function ClearImportedData(row, col){
  let activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  do {
    row = row + 1;
    var val1 = activesheet.getRange(col + row).getValue();
    var val2 = activesheet.getRange(col + (row + 1)).getValue();
  } while (val1 != "" || val2 != "")

  activesheet.getRange("A1:M" + row).clear();
}
