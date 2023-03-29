function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HWCG')
      .addItem('Reconcile', 'Reconcile')
      .addItem('Merge Payments', 'MergePayments')
      .addSeparator()
      .addItem('Format Payments', 'FormatPayments')
      .addToUi();
}

function buildInvoicePaymentMap(sheet) {
  var r = 2;
  var invoices = new Map();

  while(!sheet.getRange("A" + r).isBlank()){
    var invoiceid = sheet.getRange("B" + r).getValue();
    var type = sheet.getRange("D" + r).getValue();

    if (invoices.has (invoiceid)) {
      if (type == "Sale") {
        invoices.get(invoiceid).hasInvoice = "Sale" == type;
        // update sales row id
        invoices.get(invoiceid).saleID =r;
      }
      else if (type == "Payment") {
        invoices.get(invoiceid).hasPayment = "Payment" == type;
        // update payment row id
        invoices.get(invoiceid).paymentID =r;
      }
    } else {
      invoices.set(invoiceid, {hasPayment : "Payment" == type, hasInvoice : "Sale" == type, saleID : r, paymentID : r});
    }
    
    r =r + 1;
  }
  return invoices;
}

function buildPaymentMap(sheet){
  var r = 2;
  var payments = new Map();
  while(!sheet.getRange("A" + r).isBlank()){
    var invoiceids = sheet.getRange("E" + r).getValue();
    var hasMulti = invoiceids.toString().indexOf(" ");
    var payment = {
        No : sheet.getRange("A" + r).getValue(),
        Date : sheet.getRange("B" + r).getValue(),
        Amount : sheet.getRange("C" + r).getValue(),
        Method : sheet.getRange("D" + r).getValue(),
        EDIPayer : sheet.getRange("F" + r).getValue(),
        Payer : sheet.getRange("G" + r).getValue(),
      }
    if (hasMulti>0){
      const ids = invoiceids.split(" ");
      for ( id in ids){
        payments.set(ids[id],payment);
      }
    }
    else {
      payments.set(invoiceids,payment);
    }
    r =r + 1;
  }
  return payments;
}

function Reconcile() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var invoices= buildInvoicePaymentMap(sheet);

  for (let invoice of invoices.values()){
    if(!invoice.hasInvoice || !invoice.hasPayment){
      sheet.getRange("D" + invoice.saleID).setBackgroundRGB(255,140,140,);
    }
  }
}

function MergePayments(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDetails = ss.getSheets()[0];
  var sheetPayments = ss.getSheets()[1];
  
  // build the map
  var invoices= buildInvoicePaymentMap(sheetDetails);
  var payments = buildPaymentMap(sheetPayments);

  for ( let [id,payment] of payments.entries()){
    if (invoices.has(parseInt(id))) {
      var invoice = invoices.get(parseInt(id));
      sheetDetails.getRange("D" + invoice.saleID).setBackgroundRGB(40,160,20);
      sheetDetails.getRange("J" + invoice.paymentID).setValue(payment.No);
      sheetDetails.getRange("K" + invoice.paymentID).setValue(payment.Date);
      sheetDetails.getRange("L" + invoice.paymentID).setValue(payment.Amount);
      sheetDetails.getRange("M" + invoice.paymentID).setValue(payment.Method);
      sheetDetails.getRange("N" + invoice.paymentID).setValue(id);
      sheetDetails.getRange("O" + invoice.paymentID).setValue(payment.EDIPayer);
      sheetDetails.getRange("P" + invoice.paymentID).setValue(payment.Payer);
    }
  }
}

// copy to open area and move the mouse to the first item in the list
function FormatPayments(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let rf = sheet.getActiveCell().getRow(); // source first row
  let cf = sheet.getActiveCell().getColumn(); // source first column
  let rt = 2;
  var c = 0;

  while(!sheet.getRange(rf,cf).isBlank()){
    if (c > 6) {
      c=1;
      rt = rt + 1;
    }
    else {
      c = c + 1;
    }
    var val = sheet.getRange(rf,cf).getValue().toString();
    if ( c == 2) {
      let date = new Date(val.trim());
      val = date.toLocaleDateString("en-US",{month: '2-digit', day: '2-digit', year: 'numeric'})
    }
    else if ( c == 5) {
      if (val.length > 5) {
        var invoiceids = val.toString().match(/\d{6}/g);
        val = invoiceids.join(" ");
      }
    }
    sheet.getRange(rt,c).setValue(val);
    sheet.getRange("A" + rf).setValue("");
    rf = rf +1;
  }
}
