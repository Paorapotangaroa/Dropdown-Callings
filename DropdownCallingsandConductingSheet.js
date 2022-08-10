function gatherInfo() {
  //returns [opening Prayer, Closing Prayer]
  let prayers = getPrayers();
  let stakeB = getStakeB();
  //returns an array in the following form:
  //[[Date, Theme, Prelude, Opening Hymn, Sacrament Hymn, Musical Number/Rest Hymn , Closing Hymn, Pianist/Oranist, Chorister]]

  let musicArray = getMusicInfo();
  //returns an array in the following form:
  //[[Date, Topic, Speaker 1, Speaker 2, Conducting]]
  let talkArray = getTalkInfo();

  //returns an array in the following form:
  //[[activity1,activity2,activity3],[date1,date2,date3]]
  let announcements = getAnnoucments();


  //returns the following: 
  //[[Name, Calling, Organization], [Name, Calling, Organization], etc.]
  let sustainingsInfo = getSustainingInfo();
  //returns the following: 
  //[[Name, Calling, Organization], [Name, Calling, Organization], etc.]
  let releaseInfo = getReleaseInfo();

  let newConductingSheet = DocumentApp.create("" + musicArray[0][0].toString().substring(3, musicArray[0][0].toString().indexOf(":") - 2) + "Sacrament Conducting Sheet");
  createHead(musicArray, newConductingSheet);
  createPresiding(newConductingSheet);
  createOpeningMusicPrayer(newConductingSheet, musicArray, prayers[0]);
  createStakeB(stakeB, newConductingSheet);
  wardBusiness(newConductingSheet, releaseInfo, sustainingsInfo);
  createSacramentSection(newConductingSheet, musicArray);
  createTalkSection(newConductingSheet, talkArray, musicArray);
  createClosingSection(newConductingSheet, musicArray, prayers);
  createAnnouncements(newConductingSheet, announcements);
  //Just need announcments. Figure that out and this is done
  //for testing
  // Logger.log(musicArray);
  // Logger.log(talkArray);
  // Logger.log(sustainingsInfo);
  // Logger.log(releaseInfo);
}

function createAnnouncements(newConductingSheet, announcements) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  for(let i = 0; i<announcements[0].length; i++)
  {
    body.appendListItem(announcements[1][i].toString().substring(0,announcements[1][i].toString().indexOf(":") - 2)+" " +announcements[0][i].toString());
  }
}
function createClosingSection(newConductingSheet, musicInfo, prayers) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  let p1 = body.appendParagraph("We will close our meeting by singing a Closing Hymn: " + musicInfo[0][6] + "\n\n" +
    "After which, our Closing prayer will be offered by: " + prayers[1] + "\n\n\nAnnouncements:\n");
  p1.setAttributes(boldStyle);
}

function createTalkSection(newConductingSheet, talkInfo, musicInfo) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  let boldSection = body.appendParagraph("Speaker #1: " + talkInfo[0][2] + "\n\n" +
    "Intermediate Hymn/Musical Number: " + musicInfo[0][5] + "\n\n" +
    "Concluding Speaker: " + talkInfo[0][3] + "\n");
  boldSection.setAttributes(boldStyle);
  let everythingElse = body.appendParagraph("Expression of appreciation or added emphasis [Discretely invite comments from one who presides, as 1st option]");
  everythingElse.setAttributes(everythingElseStyle);
}

function createSacramentSection(newConductingSheet, musicInfo) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();

  let boldSection = body.appendParagraph("Sacrament (The center of our worship): We will now prepare for the sacrament by singing: \n\nHymn Number: " + musicInfo[0][4] + "\n");
  boldSection.setAttributes(boldStyle);

  let everythingElseSection = body.appendParagraph("Following the singing, the sacrament will be administered to the congregation by holders of the priesthood.\n\n\nThank you for your reverence. We also thank the priesthood holders for offering the sacrament to us. [Dismiss priesthood holders and music people to sit with members of the congregation].\n");
  everythingElseSection.setAttributes(everythingElseStyle);
}

function createStakeB(stakeB, newConductingSheet) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var body;
  var stakeBn;
  if (stakeB) {
    body = newConductingSheet.getBody();
    stakeBn = body.appendParagraph("\n\nStake Business: YES Brother Noot (said, “NOTE”)");
  }
  else {
    body = newConductingSheet.getBody();
    stakeBn = body.appendParagraph("\n\nStake Business: NO");
  }
  stakeBn.setAttributes(boldStyle);
}

function getStakeB() {
  let business;
  let ui = SpreadsheetApp.getUi();
  buttonSet = ui.ButtonSet.YES_NO;
  let response = ui.prompt("Do we have stake Stake Business? Type: \"Yes\" or \"No\"");
  if (response.getResponseText().toLowerCase() === "yes") {
    business = true;
  }
  else {
    business = false;
  }
  return business;
}

function outputReleases(releaseInfo, newConductingSheet) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  let heading = body.appendParagraph("-Releases: The following individual(s) has/have been released:");
  heading.setAttributes(boldStyle);

  for (let i = 0; i < releaseInfo.length; i++) {
    if (releaseInfo[i][2] === "Relief Society") {

    } else {
      let listItem = body.appendListItem(releaseInfo[i][0] + " as " + releaseInfo[i][1] + " in the " + releaseInfo[i][2] + "\n");
      listItem.setAttributes(everythingElseStyle);
    }
  }

  let everythingElse = body.appendParagraph("All who wish to join us in expressing appreciation for their dedicated service may do so with the uplifted hand.\n");
  everythingElse.setAttributes(everythingElseStyle);

}

function outputSustain(sustainingsInfo, newConductingSheet) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  let heading = body.appendParagraph("-Sustaining: The following individual(s) has/have been called, and we ask that you please stand and remain standing until after you are sustained:");
  heading.setAttributes(boldStyle);

  for (let i = 0; i < sustainingsInfo.length; i++) {
    if (sustainingsInfo[i][2] === "Relief Society") {

    } else {
      let listItem = body.appendListItem(sustainingsInfo[i][0] + " as " + sustainingsInfo[i][1] + " in the " + sustainingsInfo[i][2] + "\n");
      listItem.setAttributes(everythingElseStyle);
    }
  }

  let everythingElse = body.appendParagraph("\nWe propose that [he/she/they] be sustained. Those in favor, please manifest it. [Pause for vote.] Any opposed may manifest it. [Pause for vote.] Thank you.  [We invite any who are opposed to contact the bishop]\n");
  everythingElse.setAttributes(everythingElseStyle);

}

function wardBusiness(newConductingSheet, releaseInfo, sustainingsInfo) {

  //maskSection(newConductingSheet);
  propheticPromise(newConductingSheet);
  outputReleases(releaseInfo, newConductingSheet);
  outputSustain(sustainingsInfo, newConductingSheet);

}

function propheticPromise(newConductingSheet) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt("What is the prophetic promise this week?");
  let body = newConductingSheet.getBody();
  let heading = body.appendParagraph("\nProphetic Promise:");
  let p = body.appendParagraph(response.getResponseText() + "\n");
  heading.setAttributes(boldStyle)
  p.setAttributes(everythingElseStyle);
}

function maskSection(newConductingSheet) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  let bold = body.appendParagraph("Ward Business:\n\n Face Masks:");
  bold.setAttributes(boldStyle);
  let masks1 = body.appendListItem("For Sacrament meetings we, The First Presidency, encourage wearing masks when proper distancing is not possible for members in the congregation and on the rostrum.");
  let masks2 = body.appendListItem("Face masks are important for second hour meetings or council-type meetings where members are all packed in together.");
  masks1.setAttributes(everythingElseStyle);
  masks2.setAttributes(everythingElseStyle);
}

function createHead(musicArray, newConductingSheet) {
  var headingStyle = {};
  headingStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.CENTER;
  headingStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  headingStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
  headingStyle[DocumentApp.Attribute.BOLD] = true;


  let body = newConductingSheet.getBody();
  body.appendHorizontalRule();
  let firstParagraph = body.appendParagraph("Provo YSA 218th Ward Sacrament Meeting" + " - " + musicArray[0][0].toString().substring(3, musicArray[0][0].toString().indexOf(":") - 2));
  body.appendHorizontalRule();
  firstParagraph.setAttributes(headingStyle);
}

function createPresiding(newConductingSheet) {
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  let greetingAndWelcome = body.appendParagraph("Greeting & Welcome:");
  greetingAndWelcome.setAttributes(boldStyle);
  let everythingElse = body.appendParagraph("We welcome you (sisters & brothers) to our sacrament services " +
    "this afternoon. We are grateful we can worship together!\n\n - I am Bishop Barnes, and I will" +
    " conduct this meeting.\n - I am Brother Hoopes, and Bishop Barnes has asked that I conduct this meeting.")

  let presiding = body.appendParagraph("\nRecognize:\n\nPresiding Authorities:")
  let presidingInfo = body.appendParagraph(" - High Councilor: Brother Noot	");
  everythingElse.setAttributes(everythingElseStyle)
  presiding.setAttributes(boldStyle);
  presidingInfo.setAttributes(everythingElseStyle)
}

function createOpeningMusicPrayer(newConductingSheet, musicArray, prayer) {
  //   Opening Hymn:	  		#6 Redeemer of Israel		 	
  // ____Kayla Robertson____ will be our Chorister
  // ____Keola Quereto__________ will be our Pianist
  // Invocation: Following the singing, Brother Dakota Waters (Backup Clint Flinders), has been invited to give the opening prayer. 
  var boldStyle = {};
  boldStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  boldStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  boldStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  boldStyle[DocumentApp.Attribute.BOLD] = true;
  var everythingElseStyle = {};
  everythingElseStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
    DocumentApp.HorizontalAlignment.LEFT;
  everythingElseStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Verdana';
  everythingElseStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
  everythingElseStyle[DocumentApp.Attribute.BOLD] = false;
  let body = newConductingSheet.getBody();
  let p = body.appendParagraph("Opening Hymn: " + musicArray[0][3].toString() + "\n\n"
    + musicArray[0][8].toString() + " will be our Chorister" + "\n" +
    musicArray[0][7].toString() + " will be our Pianist/Organist\n");
  p.setAttributes(boldStyle);
  let p1 = body.appendParagraph("Invocation:");
  p1.setAttributes(boldStyle);
  let openingPrayerP = body.appendParagraph("Following the singing, " + prayer + " has been invited to give the opening prayer.");
  openingPrayerP.setAttributes(everythingElseStyle);

}

function getPrayers() {
  //gather prayers
  let ui = SpreadsheetApp.getUi();
  let openingPrayer = ui.prompt("Who is saying the opening prayer?");
  let closingPrayer = ui.prompt("Who is saying the closing prayer?");
  openingPrayer = openingPrayer.getResponseText();
  closingPrayer = closingPrayer.getResponseText();

  return [openingPrayer, closingPrayer];
}

//Things that vary on the conducting sheet in order of appearance. x = have already figured out how to get the correct info.
/*
  Date - x
  Presiding Authority
  Opening Hymn - x
  Chorister - x
  Pianist - x
  Opening Prayer - x
  Releases - x
  Sustaining - x
  New move ins
  Sacrament Hymn - x
  Speaker 1 - x
  Speaker 2 - x
  Intermediate Hymn/Musical Number - x
  Concluding Speaker - x
  Closing Hymn - x
  Closing Prayer - x
  Announcements
 */
function getAnnoucments() {
  let returnArray = [[], []]
  let wardCalendar = CalendarApp.getCalendarById("joqshmurs74lqgarbas6vc0t5s@group.calendar.google.com");
  let today = new Date();
  let monthfromNow = new Date();
  monthfromNow.setMonth(today.getMonth() + 1)
  events = wardCalendar.getEvents(today, monthfromNow);
  for (let i = 0; i < events.length; i++) {
    returnArray[0].push(events[i].getTitle());
    returnArray[1].push(events[i].getStartTime());
  }
  return returnArray;
}

function getSustainingInfo() {
  let sustainingsSpreadsheet = SpreadsheetApp.openById("1fjmJfi0bdtiYrsO4H9E3pVjMvVJ7H6cUSGNtxoBAe8M");
  let currentSustainingsSheet = sustainingsSpreadsheet.getSheetByName("Home");
  let roughData = currentSustainingsSheet.getSheetValues(dataStart,2,(dataLength-dataStart),7)
  let sustainingsInfoArray = [];
for(let i = 0; i<roughData.length; i++){
    if(roughData[i][4].toString()=="Called"){
      let output = [roughData[i][3],roughData[i][2],roughData[i][1]]
      sustainingsInfoArray.push(output)
    }
  }
  
  return sustainingsInfoArray;
}

function getReleaseInfo() {
  let sustainingsSpreadsheet = SpreadsheetApp.openById("1fjmJfi0bdtiYrsO4H9E3pVjMvVJ7H6cUSGNtxoBAe8M");
  let currentSustainingsSheet = sustainingsSpreadsheet.getSheetByName("Home");
  let roughData = currentSustainingsSheet.getSheetValues(dataStart,2,(dataLength-dataStart),7)
  let sustainingsInfoArray = [];
for(let i = 0; i<roughData.length; i++){
    if(roughData[i][4].toString()=="Release Extended"){
      let output = [roughData[i][3],roughData[i][2],roughData[i][1]]
      sustainingsInfoArray.push(output)
    }
  }
  
  return sustainingsInfoArray;
}

function getMusicInfo() {
  //our ward's music spreadsheet
  let musicSpreadsheet = SpreadsheetApp.openById("1sOSNnmZFA00iYlc57KNXq2kkDBOgdai20hfher4W0jg");
  let currentMusicSheet = musicSpreadsheet.getSheetByName("2022");
  let musicRoughInfoArray = currentMusicSheet.getRange(2, 1, 1000, 9).getValues();
  let currentDate = new Date(currentMusicSheet.getRange(1, 26, 1, 1).getValue());
  let musicInfoArray = [];

  for (let i = 0; i < musicRoughInfoArray.length; i++) {
    if (musicRoughInfoArray[i][0] != "") {
      //only compare dates if there is actually a date cell created
      let musicDate = new Date(musicRoughInfoArray[i][0]);
      let date1 = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());
      let date2 = new Date(musicDate.getFullYear(), musicDate.getMonth(), musicDate.getDate());

      // To calculate the time difference of two dates
      let Difference_In_Time = date2.getTime() - date1.getTime();

      // To calculate the no. of days between two dates
      let Difference_In_Days = Difference_In_Time / (1000 * 3600 * 24);

      //Make sure it is within 6 days and not in the past
      if (Difference_In_Days <= 6 && Difference_In_Days > -1) {
        musicInfoArray.push(musicRoughInfoArray[i]);
      }
    }
  }
  return musicInfoArray;
}

function getTalkInfo() {
  let talkSpreadsheet = SpreadsheetApp.openById("1M_YqOEBTSsfutXdUcMuNDugDvt4yEBHMRUH5sF6IdOk");
  let currentSheet = talkSpreadsheet.getSheetByName("Talks");
  let talkRoughInfoArray = currentSheet.getRange(2, 1, 1000, 5).getValues();
  let currentDate = new Date(currentSheet.getRange(1, 26, 1, 1).getValue());
  let talkInfoArray = [];

  for (let i = 0; i < talkRoughInfoArray.length; i++) {
    if (talkRoughInfoArray[i][0] != "") {
      //only compare dates if there is actually a date cell created
      let talkDate = new Date(talkRoughInfoArray[i][0]);
      let date1 = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate());
      let date2 = new Date(talkDate.getFullYear(), talkDate.getMonth(), talkDate.getDate());

      // To calculate the time difference of two dates
      let Difference_In_Time = date2.getTime() - date1.getTime();

      // To calculate the no. of days between two dates
      let Difference_In_Days = Difference_In_Time / (1000 * 3600 * 24);
    
      //Make sure it is within 6 days and not in the past
      if (Difference_In_Days <= 6 && Difference_In_Days > -1) {
        talkInfoArray.push(talkRoughInfoArray[i]);
      }
    }
  }
  Logger.log(talkInfoArray)
  return talkInfoArray;
}
