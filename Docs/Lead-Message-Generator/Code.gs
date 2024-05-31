function onOpen() {
  DocumentApp.getUi() // Or SpreadsheetApp or SlidesApp or FormApp.
      .createMenu('HWCG')
      .addItem('Generate Message', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Generator Lead Message');
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

function processForm(data) {
  var doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  if (data.type === "Email")
    switch (data.counter){
      case "First":
        firstEmail(body, data);
        break;
      case "Second":
        secondEmail(body, data);
        break;
      default:
        body.setText("No match " + data.type + " " + data.counter);
    }
    // body.setText(firstEmail(body, data));
  else
    switch (data.counter){
      case "First":
        firstText(body, data);
        break;
      case "Second":
        secondText(body, data);
        break;
      case "Third":
        thirdText(body, data);
        break;
      default:
        body.setText("No match " + data.type + " " + data.counter);
    }
}

function firstEmail(body, data){
 var email1 = "Hi {{Client_Name}},"
 
 var email2 = "I just left you a voicemail to discuss caregiving needs for {{Name_of_Individual_Needing_Care}}. At Homewatch CareGivers, we have spent decades offering personalized care to individuals of all ages across the country.  "
 
 var email3 ="Homewatch CareGivers has more than 250 locations across the US and 7 other countries. Our caregivers are trained, professional, experienced and have gone through extensive background and drug testing. It would be our privilege to assist {{Name_of_Individual_Needing_Care}} with {{his_her}} specific requirements.";
 
 var email4 = "I’ve attached a few items explaining who we are, what service we provide and why choose us? Also don’t forget to check our reviews on Google and Facebook. Please call us at (817) 385-6040 to schedule a no-obligation consultation. We look forward to serving {{Name_of_Individual_Needing_Care}}.";
 
 var footer="Caring expert. Expert care.";

 var his_her = "his";
 if (data.sex === "Female"){
  his_her = "her"
 }

 email1 = email1.replaceAll("{{Client_Name}}", data.caller);
 if (data.self) {
  email2 = email2.replaceAll("{{Name_of_Individual_Needing_Care}}", "yourself");
  email3 = email3.replaceAll("{{Name_of_Individual_Needing_Care}}", "you");
  email4 = email4.replaceAll("{{Name_of_Individual_Needing_Care}}", "you");
  his_her = "your";
 }
 else {
  email2 = email2.replaceAll("{{Name_of_Individual_Needing_Care}}", data.client);
  email3 = email3.replaceAll("{{Name_of_Individual_Needing_Care}}", data.client);
  email4 = email4.replaceAll("{{Name_of_Individual_Needing_Care}}", data.client);
 }
  email3 = email3.replaceAll("{{his_her}}", his_her);

 body.setText(email1);
 body.appendParagraph("");
 body.appendParagraph(email2);
 body.appendParagraph(email3);
 body.appendParagraph("");
 body.appendParagraph(email4);
 body.appendParagraph("");
 body.appendParagraph(footer);
}

function secondEmail(body, data){
  var email1="Hi {{Client_Name}},"
  
  var email2 = "This is {{Your_Name}} from Homewatch Caregivers of Irving. I am sending you this email to see if we can assist {{Name_of_Individual_Needing_Care}}. I’ve attached a few items explaining who we are, what services we provide, and why choose us? Also, here are some differentiators on why Homewatch CareGivers is the better choice versus competitors. Also don’t forget to check our reviews on Google and Facebook. We look forward to serving your family."
  
  var bullets = [
    "Annual extensive background and drug tests are mandatory for all caregivers.",
    "All caregivers are our employees, and they are not contractors. Also, they are enrolled in mandatory training every month, assigned by Homewatch Caregivers.",
    "Real person answering phones 24x7x365.",
    "Unscheduled Supervisory visit every 2 weeks.",
    "Custom care plan unique to individual needs.",
    "Quality Assurance visit from our Nurse - every 60 days.",
    "Level -3 contingency Back-up plan if caregivers call off.",
    "Bilingual caregivers are available if you prefer.",
    "No annual or monthly contracts – all we need is 2 business days advance notice for any change in the service.",
    "Access to the family portal via Kantime App on your phone to track caregiving activities.",
    "All caregivers are bonded.",
  ]

  var footer="Caring expert. Expert care.";

  email1 = email1.replaceAll("{{Client_Name}}", data.caller);
   if (data.self)
    email2 = email2.replaceAll("{{Name_of_Individual_Needing_Care}}", "you");
   else
    email2 = email2.replaceAll("{{Name_of_Individual_Needing_Care}}", data.client);

  email2 = email2.replaceAll("{{Your_Name}}", data.hwcg);

  // body.setText(email1 + '\n\n' + email2 + '\n\n' + bullets.join('\n') + '\n\n' + footer);
  body.setText(email1);
  body.appendParagraph("");
  body.appendParagraph(email2)
  var element = body.appendParagraph("");
  var ixElement = body.getChildIndex(element);
  for (bullet in bullets) {
    var listItem = body.insertListItem(ixElement + 1 + + parseInt(bullet), bullets[bullet]);
    listItem.setNestingLevel(0);
    listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
  }
  body.appendParagraph("");
  body.appendParagraph(footer);
}

function firstText(body, data){
  var text1 = "Hi {{Client_Name}}. This is {{Your_Name}} from Homewatch CareGivers of Irving. I just left you a voicemail to discuss the caregiving needs for {{Name_of_Individual_Needing_Care}}. I also sent you information via email about who we are, what service we provide and why chose us.  Please call us at (817) 385-6040 and we will be happy to discuss our services, rates and answer any other questions you may have."
  var text2 = "Have a great day."
  var text3 = "Caring experts. Expert care!!"
  var text4 = "{{Your_Name}} from Homewatch CareGivers"

  text1 = text1.replaceAll("{{Client_Name}}", data.caller);
  text1 = text1.replaceAll("{{Name_of_Individual_Needing_Care}}", data.client);
  text1 = text1.replaceAll("{{Your_Name}}", data.hwcg);

  text4 = text4.replaceAll("{{Your_Name}}", data.hwcg);

  body.setText( text1 + '\n\n' + text2 + '\n\n' + text4 + '\n\n' + text3);
}

function secondText(body, data){
  
  body.setText("No ready yet");
}

function thirdText(body, data){
  body.setText("No ready yet");
}

function fourthText(body, data){
  body.setText("No ready yet");
}
