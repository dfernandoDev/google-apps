function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HWCG')
      .addItem('Kantime to OnPay', 'Kantime2OnPay')
      .addItem('Summerize Hours', 'CreateHoursTable')
      .addItem('Summerize Mileage', 'CreateMileageTable')
      .addSeparator()
      .addItem('Clear Kantime Table', 'CleanKantimeSheet')
      .addItem('Clear OnPay Table', 'CleanOnPaySheet')
      .addToUi();
}

function Kantime2OnPay() {
  CleanOnPaySheet();
  let employees = ReadKantimeHours();
  Save2OnPay (employees);
}

function ReadKantimeHours(){
  let sheetKantime = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kantime");

  let employees = new Map();
  let r = 2;
  while(!sheetKantime.getRange("A" + r).isBlank()){
    let mileage = Number(sheetKantime.getRange("L" + r).getValue());
    let earningsCode = sheetKantime.getRange("E" + r).getValue();
    let cg = {};

    if (earningsCode == "REG") {
      cg = {
        Name : sheetKantime.getRange("A" + r).getValue(),
        ID : (sheetKantime.getRange("B" + r).getValue()),
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
        ID : (sheetKantime.getRange("B" + r).getValue()),
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

  return employees;
}

function Save2OnPay(employees){
    let sheetOnPay = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OnPay");
    let r = 2;

    let totals = new Map();
    totals = { Hours : 0, Miles : 0, Mileage: {}};

    for (const cg of employees.entries()){
      for (let attr in cg[1]) {
        if (!isNaN(parseInt(attr))) {
          let id = '00' + Number(cg[1].ID);
          sheetOnPay.getRange("A" + r).setValue(1);
          sheetOnPay.getRange("B" + r).setValue(1);
          sheetOnPay.getRange("C" + r).setValue("'" + id);
          sheetOnPay.getRange("D" + r).setValue(cg[1][attr].Hours);
          sheetOnPay.getRange("E" + r).setValue(attr);
          sheetOnPay.getRange("H" + r).setValue(cg[1].Name);
          r++;
          totals.Hours += cg[1][attr].Hours;
          if (totals.hasOwnProperty(attr)){
            totals[attr].Hours += cg[1][attr].Hours;
          } else {
            totals[attr] = {Hours : cg[1][attr].Hours};
          }
        }
      }
      if (cg[1].Mileage > 0) {
        let id = '00' + Number(cg[1].ID);
        sheetOnPay.getRange("A" + r).setValue(1);
        // sheetOnPay.getRange("B" + r).setValue(107);
        sheetOnPay.getRange("B" + r).setValue(350);
        sheetOnPay.getRange("C" + r).setValue("'" + id);
        sheetOnPay.getRange("F" + r).setValue(1);
        sheetOnPay.getRange("G" + r).setValue(cg[1].Mileage);
        sheetOnPay.getRange("H" + r).setValue(cg[1].Name);
        r++;
        // keep totals
        totals.Mileage[cg[1].ID] = {Name : cg[1].Name, Mileage : cg[1].Mileage};
        totals.Miles += cg[1].Mileage;
      }
    }
    sheetOnPay.getRange("D" + r).setFormula("=sum(D2:D" + (r - 1) + ")");
    sheetOnPay.getRange("G" + r).setFormula("=sum(G2:G" + (r - 1) + ")");

    r +=2;
    sheetOnPay.getRange("C" + r).setValue("Total Mileage");
    sheetOnPay.getRange("D" + r).setValue(totals.Miles);

    r +=2;
    sheetOnPay.getRange("C" + r).setValue("Total Hours");
    sheetOnPay.getRange("D" + r).setValue(totals.Hours);

    for (let attr in totals) {
        if (!isNaN(parseInt(attr))) {
          r++;
          sheetOnPay.getRange("C" + r).setValue(attr);
          sheetOnPay.getRange("D" + r).setValue(totals[attr].Hours);
        }
    }
}

function CleanKantimeSheet(){
    let sheetKantime = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kantime");
    let r = sheetKantime.getLastRow();

    if (r > 1) {
      sheetKantime.getRange("A2:L" + r).clear();
    }
}

function CleanOnPaySheet(){
    let sheetOnPay = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("OnPay");
    let r = sheetOnPay.getLastRow();

    if (r > 1) {
      sheetOnPay.getRange("A2:H" + r).clear();
    }
}

function CreateHoursTable() {
  let sheetKantime = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kantime");
  var sourceData = sheetKantime.getRange('A1:L24');
  var pivotTable = sheetKantime.getRange('A27').createPivotTable(sourceData);
  var pivotGroup = pivotTable.addRowGroup(1);
  pivotGroup = pivotTable.addRowGroup(2);
  pivotGroup.showTotals(false);
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['REG'])
  .build();
  pivotTable.addFilter(5, criteria);
  var pivotValue = pivotTable.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addColumnGroup(6);
};

function CreateMileageTable() {
  let sheetKantime = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kantime");
  var sourceData = sheetKantime.getRange('A1:L24');
  var pivotTable = sheetKantime.getRange('A38').createPivotTable(sourceData);
  var pivotGroup = pivotTable.addRowGroup(1);
  pivotGroup = pivotTable.addRowGroup(2);
  pivotGroup.showTotals(false);
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['MLG'])
  .build();
  pivotTable.addFilter(5, criteria);
  var pivotValue = pivotTable.addPivotValue(12, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addColumnGroup(6);
};
