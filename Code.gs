/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Build Questions', 'showSidebar')
      .addToUi();

  var ui = DocumentApp.getUi();
  var menu = ui.createMenu('Practice Multiplication')
  var subMenuIndividual = DocumentApp.getUi().createMenu('Individual Table');
  subMenuIndividual.addItem('2s', 'build2sTable');
  subMenuIndividual.addItem('3s', 'build3sTable');
  subMenuIndividual.addItem('4s', 'build4sTable');
  subMenuIndividual.addItem('5s', 'build5sTable');
  subMenuIndividual.addItem('6s', 'build6sTable');
  subMenuIndividual.addItem('7s', 'build7sTable');
  subMenuIndividual.addItem('8s', 'build8sTable');
  subMenuIndividual.addItem('9s', 'build9sTable');
  subMenuIndividual.addItem('10s', 'build10sTable');
  subMenuIndividual.addItem('11s', 'build11sTable');
  subMenuIndividual.addItem('12s', 'build12sTable');
  menu.addSubMenu(subMenuIndividual);

  var subMenuRange = DocumentApp.getUi().createMenu('Table Ranges');
  subMenuRange.addItem('2 To 3', 'build2to3Table');
  subMenuRange.addItem('2 To 4', 'build2to4Table');
  subMenuRange.addItem('2 To 5', 'build2to5Table');
  subMenuRange.addItem('2 To 6', 'build2to6Table');
  subMenuRange.addItem('2 To 7', 'build2to7Table');
  subMenuRange.addItem('2 To 8', 'build2to8Table');
  subMenuRange.addItem('2 To 9', 'build2to9Table');
  subMenuRange.addItem('2 To 10', 'build2to10Table');
  subMenuRange.addItem('2 To 11', 'build2to11Table');
  subMenuRange.addItem('2 To 12', 'build2to12Table');
  menu.addSubMenu(subMenuRange);

  menu.addSeparator();
  menu.addItem('Custom Range', 'showSidebar')
  menu.addSeparator();
  menu.addItem('Check Work', 'checkWork')
  menu.addToUi();
}

function build2sTable() {
  buildMultiplicationTable(2,2,50);
}
function build3sTable() {
  buildMultiplicationTable(3,3,50);
}
function build4sTable() {
  buildMultiplicationTable(4,4,50);
}
function build5sTable() {
  buildMultiplicationTable(5,5,50);
}
function build6sTable() {
  buildMultiplicationTable(6,6,50);
}
function build7sTable() {
  buildMultiplicationTable(7,7,50);
}
function build8sTable() {
  buildMultiplicationTable(8,8,50);
}
function build9sTable() {
  buildMultiplicationTable(9,9,50);
}
function build10sTable() {
  buildMultiplicationTable(10,10,50);
}
function build11sTable() {
  buildMultiplicationTable(11,11,50);
}
function build12sTable() {
  buildMultiplicationTable(12,12,50);
}

function build2to3Table() {
  buildMultiplicationTable(2,3,50);
}
function build2to4Table() {
  buildMultiplicationTable(2,4,50);
}
function build2to5Table() {
  buildMultiplicationTable(2,5,50);
}
function build2to6Table() {
  buildMultiplicationTable(2,6,50);
}
function build2to7Table() {
  buildMultiplicationTable(2,7,50);
}
function build2to8Table() {
  buildMultiplicationTable(2,8,50);
}
function build2to9Table() {
  buildMultiplicationTable(2,9,50);
}
function build2to10Table() {
  buildMultiplicationTable(2,10,50);
}
function build2to11Table() {
  buildMultiplicationTable(2,11,50);
}
function build2to12Table() {
  buildMultiplicationTable(2,12,50);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  //var ui = HtmlService.createHtmlOutputFromFile('sidebar')
  //    .setTitle('Range Selector');
  // DocumentApp.getUi().showSidebar(ui);

  var html = HtmlService.createTemplateFromFile("sidebar")
    .evaluate()
    .setTitle('Range Selector');

  DocumentApp.getUi().showSidebar(html);
}

function getPreferences() {
  // var tableStart = 2;
  // var tableEnd = 12;
  // var noOfQuestions = 50;

  //if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    var properties = PropertiesService.getDocumentProperties();
    var tableStart = properties.getProperty('tableStart');
    var tableEnd = properties.getProperty('tableEnd');
    var noOfQuestions = properties.getProperty('noOfQuestions');

  //}

  return {
    tableStart: userProperties.getProperty('originLang'),
    destLang: userProperties.getProperty('destLang')
  };
}

function setPreferences(tableStart, tableEnd, noOfQuestions) {
  PropertiesService.getUserProperties()
        .setProperty('tableStart', tableStart)
        .setProperty('tableEnd', tableEnd)
        .setProperty('noOfQuestions', noOfQuestions);
}

function buildMultiplicationTable(tstart, tend, noOfQuestions) {
  // Logger.log("Values s:" + tstart + " e:" + tend + " q:" + noOfQuestions);
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  doc.getHeader().clear();
  body.clear();
  var para = body.getChild(0);
  para.asParagraph().clear();

  var questions = BuildQuestions (tstart,tend, noOfQuestions);
  for (var i = 0; i < questions.length; i++) {
    para.appendText("[" + (i + 1) + "] " + questions[i] + " \n");
    //para.appendText(questions[i] + "\n");
    //body.appendParagraph(i + 1 + ") " + questions[i]);
  }
  para.setLineSpacing(2.1);
  var text = body.editAsText();
  text.setFontSize(15);
}

function checkWork() {
  var body = DocumentApp.getActiveDocument().getBody();
  var para = body.getChild(0);

  var answers = para.getText().split('\r')
  var offset = 0;
  var total = 0, missed = 0;
  for (var i = 0; i < answers.length-1; i++) {
  //for (var i = 0; i < 10; i++) {
    var validate = isCorrect(answers[i]);
    total++;
    var color = '#00AA00';
    if (!validate) {
      color = '#FF0000'
      missed++;
    }
    body.editAsText().setForegroundColor(offset, offset + answers[i].length, color);
    offset += (answers[i]).length+1;
    // console.log(answers[i] + " " + validate);
  }
  // show score
  var header = DocumentApp.getActiveDocument().getHeader();
  header.clear();
  header.editAsText().setFontSize(14);
  //header.getParagraphs()[0].setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  header.appendParagraph("Score: " + (total - missed)  + "/" + total);
}

function isCorrect(question) {
  var splitQuestion = question.split(']');
  var multiplication = (splitQuestion[1]).split('=');
  var multipliers = (multiplication[0]).split('x');
  var isValid = false;

  if (multipliers[0] * multipliers[1] == multiplication[1]) {
    isValid = true;
  }

  return isValid;
}

function BuildQuestions(tablestart, tableend, numberofquestions) {
  var table = GenerateTables(tablestart,tableend);
  var questions = GenerateQuestions(table, numberofquestions);

  return questions;
}

function GenerateTables(tablestart,tableend) {
  var table = [];
  //var numberofquestions = 10;

  // generate table for given range
  var timestable = tablestart;
  do {
    for (var i = 0; i < 12; i++) {
      table.push(timestable + " x " + (i + 1) + " =");
    }
    timestable++;
  } while (timestable <= tableend)

  // generate table for higher values but limit given range
  // this is to generate opposites 2x6 -> 6x2
  for(var i = 2; i <= 12; i++) {
    for(var j = tablestart; j <= tableend; j++) {
      if (i < tablestart || i > tableend) {
        table.push(i + " x " + j + " =");
      }
    }
  }
  return table;
}

// pick random value from the list and remove the selection
// to prevent the duplicates
function GenerateQuestions(timestable, numberofquestions) {
  // Logger.log("Generating questions.");
  var questions = [];
  for (var i = 0; i < numberofquestions && timestable.length > 0; i++) {
    var ndx = Math.floor(Math.random()*timestable.length)
    var val = timestable[ndx];
    timestable.splice(ndx, 1);
    questions.push(val);
  }
  return questions;
}

/**
 * Gets the stored user preferences for the origin and destination languages,
 * if they exist.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @return {Object} The user's origin and destination language preferences, if
 *     they exist.
 */
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  return {
    originLang: userProperties.getProperty('tableStart'),
    destLang: userProperties.getProperty('tableEnd')
  };
}
