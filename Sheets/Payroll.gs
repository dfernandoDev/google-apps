function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HWCG')
      .addItem('Convert to OnPay', 'Convert2OnPay')
      .addToUi();
}

function Convert2OnPay() {
  let sheetKantime = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kantime");
  let sheetOnPay = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OnPay");
  let employees = new Map();
  let r = 2;
  while(!sheetKantime.getRange("A" + r).isBlank()){
    let mileage = Number(sheetKantime.getRange("L" + r).getValue());
    let earningsCode = sheetKantime.getRange("E" + r).getValue();
    let cg = {};

    if (earningsCode == "REG") {
      cg = {
        Name : sheetKantime.getRange("A" + r).getValue(),
        ID : sheetKantime.getRange("B" + r).getValue(),
        Mileage : mileage,
        [sheetKantime.getRange("F" + r).getValue()] : {
          EarningsCode : earningsCode,
          Hours : sheetKantime.getRange("I" + r).getValue()
        }
      }
      if (employees.has(cg.ID)){
        let tcg = employees.get(cg.ID)
        if (tcg.hasOwnProperty(sheetKantime.getRange("F" + r).getValue())){
          tcg[sheetKantime.getRange("F" + r).getValue()].Hours += cg[sheetKantime.getRange("F" + r).getValue()].Hours;
        } else {
          tcg[sheetKantime.getRange("F" + r).getValue()] = cg[sheetKantime.getRange("F" + r).getValue()];
        }
      } else {
        employees.set(cg.ID, cg)
      }
    } else if (earningsCode == "MLG"){
      cg = {
        Name : sheetKantime.getRange("A" + r).getValue(),
        ID : sheetKantime.getRange("B" + r).getValue(),
        Mileage : mileage
      }
      if (employees.has(cg.ID)){
        employees.get(cg.ID).Mileage += mileage;
      } else {
        employees.set(cg.ID, cg)
      }
    } else {
      SpreadsheetApp.getUi().alert("Unknown Earnings Code " + earningsCode);
      break;
    }
    // console.log (hrs);
    r++;
  }
}
