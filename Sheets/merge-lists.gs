
function MergeData() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var cookbooks = new Map();

  let i = 1;
  let col = 1;
  while (!sheet.getRange(i,col).isBlank()) {
    let cookbook = sheet.getRange(i,col).getValue();
    if (cookbooks.has(cookbook)){
      Logger.log(cookbook);

    }
    else {
      cookbooks.set(cookbook,{Email : "", RVersion : "", SVersion : ""});
    }
    i = i + 1;
  }

  i = 1;
  col = 2;
  while (!sheet.getRange(i,col).isBlank()) {
    let cookbook = sheet.getRange(i,col).getValue();
    let email = sheet.getRange(i,col+1).getValue().match(/'.*'/)[0].replaceAll('\'',"");
    if (cookbooks.has(cookbook)){
      cookbooks.get(cookbook).Email = email;
    }
    else {
      Logger.log(cookbook);
    }
    i = i + 1;
  }

  i = 1;
  col = 4;
  while (!sheet.getRange(i,col).isBlank()) {
    let cookbook = sheet.getRange(i,col).getValue();
    let version = sheet.getRange(i,col+1).getValue().match(/'.*'/)[0].replaceAll('\'',"");

    if (cookbooks.has(cookbook)){
      cookbooks.get(cookbook).RVersion = version;
    }
    else {
      Logger.log(cookbook);
      cookbooks.set(cookbook,{Email : "", RVersion : version, SVersion : ""});
    }
    i = i + 1;
  }

  i = 1;
  col = 6;
  while (!sheet.getRange(i,col).isBlank()) {
    let cookbookversion = sheet.getRange(i,col).getValue();
    let arr = cookbookversion.split(" ");
    let cookbook = arr[0];
    let version = "";
    for(let a=1; a<= arr.length; a++){
      if (arr[a] != "") {
        version = arr[a];
        break;
      }

    }
    if (cookbooks.has(cookbook)){
      cookbooks.get(cookbook).SVersion = version;
    }
    else {
      Logger.log(cookbook);
      cookbooks.set(cookbook,{Email : "", RVersion : "", SVersion : version});
    }
    i = i + 1;
  }

  i = 1
  col = 7;
  cookbooks.forEach((value, key) =>{
    sheet.getRange(i,col).setValue(key);
    sheet.getRange(i,col+1).setValue(value.Email);
    sheet.getRange(i,col+2).setValue(value.RVersion);
    sheet.getRange(i,col+3).setValue(value.SVersion);
    i = i +1;
  })

}
