function geoFormatAddress() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var docNew = DocumentApp.create("Geocoded Locations");
  var bodyNew = docNew.getBody();
  var currentLines = body.getNumChildren();
  //body.appendParagraph("");
  for (let e = 0; e < currentLines; e++) {
    var company = body.getChild(e++).asText().getText();
    var address = body.getChild(e++).asText().getText();
    var city_state = body.getChild(e++).asText().getText();
    var phone = body.getChild(e++).asText().getText();
    if (phone.indexOf('+')>=0) {
      var match = phone.substring(phone.indexOf('+')+2).match(/^(\d{3})(\d{3})(\d{4})$/);
      phone = '(' + match[1] + ') ' + match[2] + '-' + match[3];
      //Logger.log(phone);
    }
    var house_no = address.substring(0,(address.indexOf(' ')));
    var street = address.substring(address.indexOf(' ')+1);
    var city = city_state.substring(0,city_state.indexOf(','));
    var no_state = city_state.substring(city_state.indexOf(' ')+1);
    var state = no_state.substring(0,no_state.indexOf(' '));
    var zip = no_state.substring(no_state.indexOf(' ')+1);
    // <street> <house number>, <zip code> <city>, <state>, <country>
    // https://docs.routexl.com/index.php/Import
    var geostring = street + ' ' + house_no + ',' + zip + ' ' + city + ',' + state + ',US,@' + company + '@' + '{' + address + '<br />' + city_state + '<br />' + phone + '}';
    //body.getChild(e).asParagraph().appendText(geostring);
    bodyNew.appendParagraph(geostring);
    //body.appendParagraph(geostring);
    Logger.log(geostring);
  }
}
