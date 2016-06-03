/**
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */

var fbURL = 'https://www.facebook.com/groups/541841935909841/';
var maestroCzarEmail = 'Marc Majcher <majcher@gmail.com>';

var maestroData = {
  col: 'B',
  colNum: 2,
  showType: '',
  callTime: '9:20pm'
};

var rawData = {
  col: 'G',
  colNum: 7,
  showType: 'RAW',
  callTime: '5:20pm'
};

/**
 * Generate message to post for casting Maestro
 */
function generateMaestroCasting() {
  generateCastingMessage(maestroData);
}

/**
 * Generate message to post for casting Maestro RAW
 */
function generateMaestroRawCasting() {
  generateCastingMessage(rawData);
}

function generateCastingMessage(data) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();

  var showDate = sheet.getName();
  var director1 = range.getCell(17, data.colNum).getValue();
  var director2 = range.getCell(18, data.colNum).getValue();
  var musician = range.getCell(15, data.colNum).getValue();
  var signup = range.getCell(22, data.colNum).getValue();

  // Bail if it's not a casting sheet
  if (range.getCell(1, 1).getValue() !== 'Maestro Cast') {
    return;
  }
  if (!director1 || !director2) {
    SpreadsheetApp.getUi().alert('Director Missing!');
    return;
  }
  if (!musician) {
    SpreadsheetApp.getUi().alert('Musician Missing!');
    return;
  }
  if (!signup) {
    SpreadsheetApp.getUi().alert('No signup URL!');
    return;
  }

  var message;
  if (data.showType === 'RAW') {
    message = 'Hello!\n\n';
    message += 'Your Maestro RAW directors this week (' + showDate + ') are ';
    message += director1 + ' and ' + director2 + '!\n\n';
    message +=
      'Sign up if you are a current or recent graduate of the Hideout Theatre, and bring your people!\n\n';
    message +=
      'Call time is 5:20. Warm-ups will be at 5:30. Show is at 6pm! (it is ill-advised to be late).\n\n';
    message += musician + ' will be our musician this week!!\n\n';
    message += 'TAKE ADVANTAGE OF THIS AMPLE STAGE TIME!!!\n\n';
    message +=
      '(Sign up with your full name and experience level via the form below. I can\'t cast you if you don\'t put your name and experience level.)\n\n';
    message += 'v v v v v SIGN UP HERE! v v v v v\n';
    message += signup + '\n';
    message += '^ ^ ^ ^ ^ SIGN UP HERE! ^ ^ ^ ^ ^\n\n'

  }
  else {
    message = 'Hello everyone!\n';
    message += '10pm Maestro - ' + showDate + '\n\n';
    message +=
      'Maestro is the Hideout\'s Saturday 10 pm!It features games and scenes and nail - biting eliminations!\n\ n ';
    message += 'Directors: ' + director1 + ' and ' + director2 + '!\n';
    message += 'Music: ' + musician + '!\n\n';
    message +=
      'Sign up via the Google Form below with your NAME, your EXPERIENCE LEVEL, and your FAVORITE WARMUP.\n\n';
    message +=
      '(You\'ll only be cast for tech if you volunteer to do tech.Also, if you tech, you \'re automatically cast in the next show!)\n\n';
    message += 'v v v v v SIGN UP HERE! v v v v v\n';
    message += signup + '\n';
    message += '^ ^ ^ ^ ^ SIGN UP HERE! ^ ^ ^ ^ ^\n\n'
    message +=
      '_ .... .. ... .. ... _. ___ _ ._ ... .__. . _._. .. ._ ._.. __ . ... ... ._ __. . \n\n';
    message += 'If you sign up, please consider the following:\n';
    message +=
      '1. Call-time is a hard 9:20pm. Please do not be late. Warm-ups will be organized and led by a director or a pre-selected member of the cast.\n';
    message +=
      '2. If you are in (or are planning on seeing) the 8pm mainstage, please mention that. I will try to only cast one or two players involved with the 8pm mainstage, since it forces them to miss call-time.\n';
    message +=
      '3. Staying for notes is required. We are working hard for Maestro to get out by 11:45 and for notes to last no longer than 15 minutes. Feel free to drink your free beer and watch the show if eliminated!\n';
    message +=
      '4. Dropping out of Maestro after getting cast results in a one month no-play penalty\n';
  }
  showMessage('Maestro' + data.showType + 'Casting Notice', message, fbURL);
}

/**
 * Generate message to announce Maestro cast
 */
function generateMaestroCast() {
  generateCastMessage(maestroData);
};

/**
 * Generate message to announce Maestro RAW cast
 */
function generateMaestroRawCast() {
  generateCastMessage(rawData);
}

function generateCastMessage(data) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();

  var showDate = sheet.getName();
  var director1 = range.getCell(17, data.colNum).getValue();
  var director2 = range.getCell(18, data.colNum).getValue();
  var musician = range.getCell(15, data.colNum).getValue();
  var tech = range.getCell(14, data.colNum).getValue();
  var postURL = range.getCell(20, data.colNum).getValue();


  // Bail if it's not a casting sheet
  if (range.getCell(1, 1).getValue() !== 'Maestro Cast') {
    return;
  }
  if (!director1 || !director2) {
    SpreadsheetApp.getUi().alert('Director Missing!');
    return;
  }
  if (!musician) {
    SpreadsheetApp.getUi().alert('Musician Missing!');
    return;
  }
  if (!tech) {
    SpreadsheetApp.getUi().alert('Tech Person Missing!');
    return;
  }

  var message = 'Your Maestro ' + data.showType + ' cast for this week, ' +
    showDate + ':\n\n';
  message += 'Directors - ' + director1 + ' and ' + director2 + '\n';
  message += 'Music by ' + musician + '\n';
  message += tech + ' on tech\n\nPlayers:\n';

  // Sort and grab the cast list
  var players = sheet.getRange(data.col + '1:' + data.col + '12')
    .sort(data.colNum).getValues();
  for (var i = 0; i < players.length; i++) {
    message += players[i][0] + '\n';
  }

  message += '\nCall time is ' + data.callTime +
    ' - come on down and ready to warm up!'
  showMessage('Maestro ' + data.showType + ' Cast Notice', message, postURL);
}

/**
 * Add sheet to spreadsheet for new week of Maestro casting
 */
function addNewWeek() {
  var sheetDate = SpreadsheetApp.getActiveSheet().getName();
  var dateArray = sheetDate.split('/');
  var month = Number(dateArray[0]) - 1;
  var day = dateArray[1];
  var year = dateArray[2];
  var date = new Date(year, month, day);
  date.setDate(date.getDate() + 7);
  var dateString = (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();

  // duplicate sheet and rename
  SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.setName(dateString);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(3);

  // clear Maestro players
  var players = sheet.getRange('B1:B12').sort(2);
  players.moveTo(sheet.getRange('B50'));
  var signups = sheet.getRange('D1:D49').sort(4);
  signups.moveTo(sheet.getRange('D50'));
  var tech = sheet.getRange('B14');
  tech.moveTo(sheet.getRange('B1'));
  sheet.getRange('B15:B21').clear();

  // clear RAW players
  var players = sheet.getRange('G1:G12').sort(7);
  players.moveTo(sheet.getRange('G50'));
  var signups = sheet.getRange('I1:I49').sort(9);
  signups.moveTo(sheet.getRange('I50'));
  var tech = sheet.getRange('G14');
  tech.moveTo(sheet.getRange('G1'));
  sheet.getRange('G15:G21').clear();

  // grab this week's directors
  var schedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    'Schedule').getDataRange().getValues();
  for (var i = 1; i < schedule.length; i++) {
    var thisDate = new Date(schedule[i][0]);
    var thisDateString = (thisDate.getMonth() + 1) + '/' + thisDate.getDate() +
      '/' + thisDate.getFullYear();
    if (thisDateString.match(dateString)) {
      sheet.getRange('B17').setValue(schedule[i][1]);
      sheet.getRange('B18').setValue(schedule[i][2]);
      sheet.getRange('G17').setValue(schedule[i][5]);
      sheet.getRange('G18').setValue(schedule[i][6]);
      sheet.getRange('B15').setValue(schedule[i][3]);
      sheet.getRange('G15').setValue(schedule[i][7]);
      break;
    }
  }

  SpreadsheetApp.setActiveSheet(sheet);
}

/**
 * Send email notifying house manager, directors, and musician for this week's Maestro
 */
function maestroReminderEmail() {
  sendReminderEmail(maestroData);
}

/**
 * Send email notifying house manager, directors, and musician for this week's Maestro RAW
 */
function maestroRawReminderEmail() {
  sendReminderEmail(rawData);
}

function sendReminderEmail(data) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var contacts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    'Contacts');

  var houseManagerEmail = contacts.getRange('B2').getValue();
  var musicianContacts = contacts.getRange('B6:B10').getValues();
  var directorContacts = contacts.getRange('A2:A50').getValues();

  var musician = sheet.getRange(data.col + '15').getValue();
  var director1 = sheet.getRange(data.col + '17').getValue();
  var director2 = sheet.getRange(data.col + '18').getValue();

  for (var i = 0; i < musicianContacts.length; i++) {
    if (musicianContacts[i][0].match(musician)) {
      var musicianEmail = musicianContacts[i][0];
      break;
    }
  }

  for (var i = 0; i < directorContacts.length; i++) {
    if (directorContacts[i][0].match(director1)) {
      var director1Email = directorContacts[i][0];
    }
    if (directorContacts[i][0].match(director2)) {
      var director2Email = directorContacts[i][0];
    }
    if (director1Email && director2Email) {
      break;
    }
  }

  var message = 'Hey ' + getFirstName(houseManagerEmail) + ', ' + getFirstName(
    director1) + ', ' + getFirstName(director2) + ', and ' + getFirstName(
    musician) + '!\n\n';
  message += getFirstName(houseManagerEmail) + ' is managing house, and ' +
    getFirstName(musician) + ' is our musician!\n\n';
  message += getFirstName(director1) + ', ' + getFirstName(director2) +
    ', call is at ' + data.callTime +
    ' - take a look at what people like for warmups, and make sure they get into them!\n\n' +
    sheet.getRange('G23').getValue() + '\n\nYour cast:\n\n';
  message += 'Tech: ' + sheet.getRange('G14').getValue() + '\n\nPlayers:\n';

  var players = sheet.getRange(data.col + '1:' + data.col + '12')
    .sort(data.colNum).getValues();
  for (var i = 0; i < players.length; i++) {
    message += players[i][0] + '\n';
  }

  message += '\n\nHave Fun!\n\nThanks,\nMarc\n';

  var recipients = [maestroCzarEmail, houseManagerEmail, director1Email,
    director2Email, musicianEmail
  ].join(',');

  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Sending to: ' + recipients, message, ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.OK) {
    GmailApp.sendEmail(recipients, 'You\'re involved in Maestro ' + data.showType +
      'on ' + sheet.getName() + '!', message);
    SpreadsheetApp.getActiveSpreadsheet().toast('Message sent.');
  }
  else {
    SpreadsheetApp.getActiveSpreadsheet().toast('Message not sent.');
  }
}

function getFirstName(fullName) {
  var nameArr = fullName.split(' ');
  return (nameArr[0]);
}

/**
 * Display message in text box with link
 */
function showMessage(title, message, url) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Add menus to spreadsheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Maestro Actions')
    .addItem('Generate Maestro Casting Notice', 'generateMaestroCasting')
    .addItem('Generate Maestro RAW Casting Notice', 'generateMaestroRawCasting')
    .addSeparator()
    .addItem('Generate Maestro Cast Message', 'generateMaestroCast')
    .addItem('Generate Maestro RAW Cast Message', 'generateMaestroRawCast')
    .addSeparator()
    .addItem('Send Maestro Reminder Email', 'maestroReminderEmail')
    .addItem('Send Maestro RAW Reminder Email', 'maestroRawReminderEmail')
    .addSeparator()
    .addItem('Create New Week', 'addNewWeek')
    .addToUi();
}
