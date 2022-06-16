function getLocationData() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var prospects = {};
  var locations = [];
  prospects.locations = locations;

  for (let e = 0; e < body.getNumChildren(); e++) {
    var location =[];  
    var company = body.getChild(e++).asText().getText();
    location['company'] = company;
    var address = body.getChild(e++).asText().getText();
    location['address'] = address;
    var city_state = body.getChild(e++).asText().getText();
    location['address'] = address;
    var phone = body.getChild(e++).asText().getText();
    if (phone.indexOf('+')>=0) {
      var match = phone.substring(phone.indexOf('+')+2).match(/^(\d{3})(\d{3})(\d{4})$/);
      phone = '(' + match[1] + ') ' + match[2] + '-' + match[3];
    }
    else {
      var match = phone.match(/^(\d{3}).(\d{3}).(\d{4})$/);
      phone = '(' + match[1] + ') ' + match[2] + '-' + match[3];
    }
    location['phone'] = phone;
    var house_no = address.substring(0,(address.indexOf(' ')));
    location['house_no'] = house_no;
    var street = address.substring(address.indexOf(' ')+1);
    location['street'] = street;
    var city = city_state.substring(0,city_state.indexOf(','));
    location['city'] = city;
    var no_state = city_state.substring(city_state.indexOf(' ')+1);
    location['no_state'] = no_state;
    var state = no_state.substring(0,no_state.indexOf(' '));
    location['state'] = state;
    var zip = no_state.substring(no_state.indexOf(' ')+1);
    location['zip'] = zip;
    prospects.locations.push(location);
  }

  return prospects;
}

function exportDoc() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var docNew = DocumentApp.create("Geocoded Locations");
  var bodyNew = docNew.getBody();
  var currentLines = body.getNumChildren();
  //body.appendParagraph("");
  var prospects = getLocationData();
  for (let location in prospects.locations) {
    // <street> <house number>, <zip code> <city>, <state>, <country>
    // https://docs.routexl.com/index.php/Import
    var geostring = prospects.locations[location].street + ' ' + prospects.locations[location].house_no + ',' + prospects.locations[location].zip + ' ' + prospects.locations[location].city 
        + ',' + prospects.locations[location].state + ',US,@' + prospects.locations[location].company + '@' 
        + '{' + prospects.locations[location].address + '<br />' + prospects.locations[location].city_state + '<br />' + prospects.locations[location].phone + '}';
    //body.getChild(e).asParagraph().appendText(geostring);
    bodyNew.appendParagraph(geostring);
    //body.appendParagraph(geostring);
    Logger.log(geostring);
  }
  bodyNew.getChild(0).removeFromParent();
}

function exportSheet() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var spreadsheet = SpreadsheetApp.create("Geo Locations");
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange("A1").setValue("Company");
  sheet.getRange("B1").setValue("Address");
  sheet.getRange("C1").setValue("City");
  sheet.getRange("D1").setValue("State");
  sheet.getRange("E1").setValue("Zip");
  sheet.getRange("F1").setValue("Phone");
  var c = 2;
  var prospects = getLocationData();
  for (let e = 0; e < prospects.locations.length; e++) {
    sheet.getRange("A" + c).setValue(prospects.locations[e].company);
    sheet.getRange("B" + c).setValue(prospects.locations[e].address);
    sheet.getRange("C" + c).setValue(prospects.locations[e].city);
    sheet.getRange("D" + c).setValue(prospects.locations[e].state);
    sheet.getRange("E" + c).setValue(prospects.locations[e].zip);
    sheet.getRange("F" + c).setValue(prospects.locations[e].phone);
    c++;
  }
}
