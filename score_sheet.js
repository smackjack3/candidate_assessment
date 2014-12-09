/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Upload to Taleo', functionName: 'uploadScores_'}
  ];
  spreadsheet.addMenu('Taleo', menuItems);
}

//-------------------------------------------------------------------------------------------------//
//  This function is called by when a user chooses the Upload to Taleo menu option.  It calls the  //
//  calcScores function and then uploads the optimism and ASQ scores, and essays, to Taleo.        //                  
//-------------------------------------------------------------------------------------------------//

function uploadScores_() {
  var ss = SpreadsheetApp.getActiveSheet();
  var values = ss.getDataRange();
  var numCols = values.getNumColumns();
  
  // Prompt the user for a row number.
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt("Enter candidate row",
                             "Please enter the row number containing the\ncandidate information you would like to upload.\n" +
                              "(For example: 2)", 
                             ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() != ui.Button.OK) {
    return;
  } else {
    var selectedRow = prompt.getResponseText();
  }
  
  var rowNumber = Number(selectedRow);
  if (isNaN(rowNumber) || rowNumber < 2 ||
      rowNumber > ss.getLastRow()) {
    ui.alert('Error',
        Utilities.formatString('Row %s is not valid.', selectedRow),
        ui.ButtonSet.OK);
    return;
  }

  var data = ss.getRange(rowNumber, 1, 1, numCols).getValues();
  var results = calcScores_(data);


  //Retrieve authToken from Taleo
  var authToken = Taleo.login();
  
  if(authToken) {
    try {
      //Post data to candidate record
      var payload = JSON.stringify({ "candidate" : { "gritScore_txt" : JSON.parse(results).grit, "asqScore_txt" : JSON.parse(results).asq, "ntfEssay1" : JSON.parse(results).essays.essay1, "ntfEssay2" : JSON.parse(results).essays.essay2, "ntfEssay3" : JSON.parse(results).essays.essay3 }});
      Taleo.submit(authToken, 'https://ch.tbe.taleo.net/CH06/ats/api/v1/object/candidate/' + JSON.parse(results).taleoId, 'PUT', payload, 'application/json; charset=utf-8');
    } 
    catch(e) {
      alert(e);
    }
    //Logout
    Taleo.logout(authToken);
    return true;
  }
  return false;
}

//-------------------------------------------------------------------------------//
//  This function is called when the form is submitted. It calls the calcScores  //
//  function and then uploads the optimism and ASQ scores, and essays, to Taleo. //                  
//-------------------------------------------------------------------------------//

function onFormSubmit() {
  var ss = SpreadsheetApp.getActiveSheet();
  var values = ss.getDataRange();

  var data = ss.getRange(values.getLastRow(), 1, 1, values.getNumColumns()).getValues();
  var results = calcScores_(data);
  
  if(checkForDups_(values, JSON.parse(results).taleoId)) {
    //Create custom error content
    var error = {
      message : "Duplicate found: " + JSON.parse(results).taleoId,
      fileName : "",
      lineNumber : ""
    };
    writeError_(error);
    return;
  }

  if(JSON.parse(results).grit >= 3.3 && JSON.parse(results).asq < 14) {
    var passed = true;
  }
  
  //Retrieve authToken from Taleo
  var authToken = Taleo.login();
  
  if(authToken) {
    try {
      //Post data to candidate record
      var payload = JSON.stringify({ "candidate" : { "gritScore_txt" : JSON.parse(results).grit, "asqScore_txt" : JSON.parse(results).asq, "ntfEssay1" : JSON.parse(results).essays.essay1, "ntfEssay2" : JSON.parse(results).essays.essay2, "ntfEssay3" : JSON.parse(results).essays.essay3 }});
      Taleo.submit(authToken, 'https://ch.tbe.taleo.net/CH06/ats/api/v1/object/candidate/' + JSON.parse(results).taleoId, 'PUT', payload, 'application/json; charset=utf-8');
      var email = getCandidateEmail_(authToken, JSON.parse(results).taleoId);
      sendEmail_(email);
      if(passed) {
        payload = JSON.stringify({ "candidateapplication" : { "candidateId" : JSON.parse(results).taleoId, "requisitionId" : 906, "status" : 5038 }});
        Taleo.submit(authToken, 'https://ch.tbe.taleo.net/CH06/ats/api/v1/object/candidateapplication?upsert=true', 'POST', payload, 'application/json; charset=utf-8');
      }
      else {
        payload =  JSON.stringify({ "candidateapplication" : { "candidateId" : JSON.parse(results).taleoId, "requisitionId" : 906, "status" : 5042 }}); /*This isn't working yet so commented out*/ //"reasonRejected": "Did not pass candidate assessment, TF" }});
        Taleo.submit(authToken, 'https://ch.tbe.taleo.net/CH06/ats/api/v1/object/candidateapplication?upsert=true', 'POST', payload, 'application/json; charset=utf-8');
        addToDeclinationQueue_(JSON.parse(results).taleoId, JSON.parse(results).firstName, JSON.parse(results).lastName, email);
      }
    } 
    catch(e) {
      writeError_(e);
    } 
    //Logout
    Taleo.logout(authToken);
    return true;
  }
  return false;
}


//-----------------------------------------------------------------------------------------------//
//  This function calculates the optimism and ASQ scores and returns them along with the essays  //
//-----------------------------------------------------------------------------------------------//

function calcScores_(values) {
  
  //Declare variables
  var gritScore = 0, asqScore = 0;
  var finalGritScore = 0.0, finalAsqScore = 0.0;
  var returnObject;
  
  //----------------------------------------------//
  //  This section will calculate the Grit score  //
  //----------------------------------------------//
  
  //Loop through 12 grit assessment items (4 text fields in beginning (timestamp, id, first name and last name) are skipped over) 
  for (var i = 4; i < 16; i++) {
    var itemResponse = values[0][i];
    if (i==4 || i==7 || i==9 || i==12 || i==13 || i==15) {
      switch (itemResponse) {
        case 'Very much like me':
          gritScore += 5;
          break;
        case 'Mostly like me':
          gritScore += 4;
          break;
        case 'Somewhat like me':
          gritScore += 3;
          break;
        case 'Not much like me':
          gritScore += 2;
          break;
        case 'Not like me at all':
          gritScore += 1;
          break;
      }
    } else if (i==5 || i==6 || i==8 || i==10 || i==11 || i==14) {
      switch (itemResponse) {
        case 'Very much like me':
          gritScore += 1;
          break;
        case 'Mostly like me':
          gritScore += 2;
          break;
        case 'Somewhat like me':
          gritScore += 3;
          break;
        case 'Not much like me':
          gritScore += 4;
          break;
        case 'Not like me at all':
          gritScore += 5;
          break;
      }
    }
  }
  
  //Return final grit score
  finalGritScore = gritScore / 12;
  
  //---------------------------------------------//
  //  This section will calculate the ASQ score  //
  //---------------------------------------------//
  
  //Loop through items 16-39 for the CoNeg questions
  for (var i = 16; i < 40; i++) {
    var itemResponse = values[0][i];
    if (!(isNaN(itemResponse))) {
      asqScore += parseInt(itemResponse);
    }
  }
  
  //Return final ASQ score
  finalAsqScore = asqScore / 6;
    
  //--------------------------------------------------------------------------------------------------//
  //  This section creates a JSON object to hold the scores and essay questions for sending to Taleo  //
  //--------------------------------------------------------------------------------------------------//
  
  return returnObject = JSON.stringify({ "taleoId" : values[0][1], "firstName" : values[0][2], "lastName" : values[0][3], "grit" : finalGritScore.toFixed(2), "asq" : finalAsqScore.toFixed(2), "essays" : { "essay1" : escapeSpecialChars_(values[0][40]), "essay2" : escapeSpecialChars_(values[0][41]), "essay3" : escapeSpecialChars_(values[0][42]) } });

}

//---------------------------------------------------------------------------------------------------------------//
//  This function replaces control characters found in the passed in string and is used for the essay responses  //
//---------------------------------------------------------------------------------------------------------------//

function escapeSpecialChars_(str) {
    return str.replace(/\\n/g, "\\n")
               .replace(/\\'/g, "\\'")
               .replace(/\\"/g, '\\"')
               .replace(/\\&/g, "\\&")
               .replace(/\\r/g, "\\r")
               .replace(/\\t/g, "\\t")
               .replace(/\\b/g, "\\b")
               .replace(/\\f/g, "\\f");
}

//-----------------------------------------------------------------------------------------//
//  This function checks for duplicate responses and returns true if duplicates are found  //
//-----------------------------------------------------------------------------------------//

function checkForDups_(values, taleoId) {
  Logger.log('Checking for previous submissions.');
  var numOfRows = values.getNumRows();
  var response = values.getValues();
  var count = 0;
  //Loop through Taleo IDs to check for a match
  for(var i = 0; i < numOfRows; i++) {
    var existingTaleoId = response[i][1];
    if(taleoId == existingTaleoId) {
      count++;
      if(count > 1) {
        Logger.log('Previous submission FOUND.');
        return true;
      }
    }
  }
  Logger.log('Previous submission NOT FOUND.');
  return false;
}

function getCandidateEmail_(authToken, taleoId) {
  if(authToken) {
    try {
      var submitURL = 'https://ch.tbe.taleo.net/CH06/ats/api/v1/object/candidate/' + taleoId;
      var response = Taleo.submit(authToken, submitURL, 'GET', null, null);
    } 
    catch(e) {
      writeError_(e);
    } 
    return JSON.parse(response).response.candidate.email;
  }
  return false; 
}

function sendEmail_(email) {
  var subject = "Citizen Schools | Confirmation - Application Part 2";
  var body = [
    "Thank you for completing part two of your application for the National Teaching Fellowship.\n\n",
    "We will be reviewing your information and will contact you within the next 1-2 weeks. In the meantime, check out what's happening in Citizen Schools by viewing our <a href=\"https://www.youtube.com/user/citizenschoolstv/featured\">CitizenSchoolsTV</a> channel on YouTube. You can also learn more about the National Teaching Fellowship by reading some posts from our <a href=\"http://www.citizenschools.org/blog/tag/teaching-fellowship/\">InspirED blog</a>.\n\n",
    "In service,\n\n",
    "Talent Acquisition\n",
    "Citizen Schools"
  ];
  var htmlBody = [];
  for(var i=0; i<body.length; i++) { 
    htmlBody.push(body[i].replace(/\n/g, "<br />")); 
  }
  
  try {
    MailApp.sendEmail(email, subject, body.join(''), {
      htmlBody: htmlBody.join(''),
      noReply: true
    });
  } catch(e) {
    writeError_(e);
  }
}

function addToDeclinationQueue_(taleoId, fName, lName, email) {
  var sheet = SpreadsheetApp.openById('[INSERT_SHEET_ID]').getSheetByName('Sheet1');
  var lastRow = sheet.getLastRow();
  sheet.getRange('A1').offset(lastRow, 0).setValue(new Date());
  sheet.getRange('A1').offset(lastRow, 1).setValue(taleoId);
  sheet.getRange('A1').offset(lastRow, 2).setValue(fName);
  sheet.getRange('A1').offset(lastRow, 3).setValue(lName);
  sheet.getRange('A1').offset(lastRow, 4).setValue(email);
}
                        

function writeError_(e) {
  var sheet = SpreadsheetApp.openById('[INSERT_SHEET_ID]').getSheetByName(SpreadsheetApp.getActive().getName());
  var lastRow = sheet.getLastRow();
  var cell = sheet.getRange('A1');
  
  cell.offset(lastRow, 0).setValue(new Date());
  cell.offset(lastRow, 1).setValue(e.message);
  cell.offset(lastRow, 2).setValue(e.fileName);
  cell.offset(lastRow, 3).setValue(e.lineNumber);
  
  MailApp.sendEmail("rorysmith@citizenschools.org", "Script Failed : NTF Application (Part 2)", "Error: " + e.message + "\nFile: " + e.fileName + "\nLine: " + e.lineNumber);
}
