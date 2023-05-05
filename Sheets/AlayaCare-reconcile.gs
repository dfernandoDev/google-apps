function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HWCG')
      .addItem('Merge Payments Bette', 'MergePaymentsBette')
      .addItem('Merge Payments Sandra', 'MergePaymentsSandra')
      .addItem('Reconcile', 'Reconcile')
      .addSeparator()
      .addItem('Format Payments', 'FormatPayments')
      .addToUi();
}

function buildInvoicePaymentMap(sheet) {
  let r = 2;
  let invoices = new Map();

  while(!sheet.getRange("A" + r).isBlank()){
    let invoiceid = sheet.getRange("B" + r).getValue();
    let type = sheet.getRange("D" + r).getValue();

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
    
    r = r + 1;
  }
  return invoices;
}

function buildPaymentMap(sheet){
  let r = 2;
  let payments = new Map();
  while(!sheet.getRange("A" + r).isBlank()){
    let invoiceids = sheet.getRange("E" + r).getValue();
    let hasMulti = invoiceids.toString().indexOf(" ");
    let payment = {
        No : sheet.getRange("A" + r).getValue(),
        Date : sheet.getRange("B" + r).getValue(),
        Amount : sheet.getRange("C" + r).getValue(),
        Method : sheet.getRange("D" + r).getValue(),
        AMS : sheet.getRange("E" + r).getValue(),
        EDIPayer : sheet.getRange("F" + r).getValue(),
        Payer : sheet.getRange("G" + r).getValue(),
        RowID : r,
      }
    if (hasMulti>0){
      const ids = invoiceids.split(" ");
      for ( id in ids){
        if (payments.has(id)){
          SpreadsheetApp.getUi().alert("Duplicate payments for invoice id: " + ids[id]);
          Logger.log(id);
        }
        else {
          payments.set(ids[id],payment);
        }
      }
    }
    else {
      payments.set(invoiceids,payment);
    }
    r =r + 1;
  }
  return payments;
}

function MergePaymentsBette(){
  let sheetACaccounting = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AC Accounting-Bette"); // ss.getSheets()[0];
  MergePayments(sheetACaccounting);
}

function MergePaymentsSandra(){
  let sheetACaccounting = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AC Accounting-Sandra"); // ss.getSheets()[0];
  MergePayments(sheetACaccounting);
}

function MergePayments(sheetACaccounting){
  let sheetPayments = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("VA Payments"); // ss.getSheets()[1];
  // build the map
  let invoicespayments= buildInvoicePaymentMap(sheetACaccounting);
  let payments = buildPaymentMap(sheetPayments);

  for ( let [id,payment] of payments.entries()){
    if (invoicespayments.has(parseInt(id))) {
      var invoice = invoicespayments.get(parseInt(id));
      // write the corresponding payment information next to sale
      sheetACaccounting.getRange("D" + invoice.saleID).setBackgroundRGB(40,160,20);
      sheetACaccounting.getRange("J" + invoice.saleID).setValue(payment.No);
      sheetACaccounting.getRange("K" + invoice.saleID).setValue(payment.Date);
      sheetACaccounting.getRange("L" + invoice.saleID).setValue(payment.Amount);
      sheetACaccounting.getRange("M" + invoice.saleID).setValue(payment.Method);
      sheetACaccounting.getRange("N" + invoice.saleID).setValue(payment.AMS);
      sheetACaccounting.getRange("O" + invoice.saleID).setValue(payment.EDIPayer);
      sheetACaccounting.getRange("P" + invoice.saleID).setValue(payment.Payer);
      // write the corresponding payment information next to payment
      sheetACaccounting.getRange("J" + invoice.paymentID).setValue(payment.No);
      sheetACaccounting.getRange("K" + invoice.paymentID).setValue(payment.Date);
      sheetACaccounting.getRange("L" + invoice.paymentID).setValue(payment.Amount);
      sheetACaccounting.getRange("M" + invoice.paymentID).setValue(payment.Method);
      sheetACaccounting.getRange("N" + invoice.paymentID).setValue(payment.AMS);
      sheetACaccounting.getRange("O" + invoice.paymentID).setValue(payment.EDIPayer);
      sheetACaccounting.getRange("P" + invoice.paymentID).setValue(payment.Payer);
    }
    else if (payment.AMS.toString().indexOf(id) == -1) {
      sheetPayments.getRange("A" + payment.RowID).setBackgroundRGB(245,155,155);
    }
  }
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
        var invoiceids = val.toString().match(/\d{6}|\d{3},\d{3}-\d/g);
        
        val = invoiceids.join(" ").replace(",","");
      }
    }
    sheet.getRange(rt,c).setValue(val);
    sheet.getRange("A" + rf).setValue("");
    rf = rf +1;
  }
}
