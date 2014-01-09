function onFormSubmit(e) {
  //Declare form ID
  var formId = '1zLerefHduVdRH-X0aEeAyS7bVUwLx7wcYvx67nG2oQE';

  //Get form responses
  try {
    var formResponse = e.response.getItemResponses();
  }
  catch(ex) {
    //Write error to error log
    var doc = DocumentApp.openById('1VCnA9TCcfIh1MyyRzyOct3UDNKEzuBJko1buYZ5vP_s');
    doc.getBody().appendParagraph('[' + new Date() + '] Error getting form responses: ' + ex);
    return false;
  }
  
  //Calculate scores
  try {
    var results = calcScores(formResponse);
  }
  catch(ex) {
    //Write error to error log
    var doc = DocumentApp.openById('1VCnA9TCcfIh1MyyRzyOct3UDNKEzuBJko1buYZ5vP_s');
    doc.getBody().appendParagraph('[' + new Date() + '] Error calculating scores: ' + ex);
    return false;
  }
   
  //Submit to Taleo
  try {
    //Check for duplicates and only submit data if none are found
    if(checkForDups(JSON.parse(results).taleoId, formId)) {
      Logger.log('Previous submission found.  Submission will not be submitted to Taleo.');
      //Write error to error log
      var doc = DocumentApp.openById('1VCnA9TCcfIh1MyyRzyOct3UDNKEzuBJko1buYZ5vP_s');
      doc.getBody().appendParagraph('[' + new Date() + '] Duplicate(s) found for Taleo ID ' + JSON.parse(results).taleoId + '. Results not submitted.');
    }
    else {
      Logger.log('Previous submission not found.');
      //Grab authToken from Taleo, submit scores, and logout
      var authToken = taleoLogin();
      taleoSubmit(authToken, results);
      taleoLogout(authToken);
    }
  }
  catch(ex) {
    //Write error to error log
    var doc = DocumentApp.openById('1VCnA9TCcfIh1MyyRzyOct3UDNKEzuBJko1buYZ5vP_s');
    doc.getBody().appendParagraph('[' + new Date() + '] Error submitting scores: ' + ex);
    return false;
  }
}


//--------------------------------------------------------------------------------------------------------------------//
//  This function scores the assessments and returns the total score, the essays, and the Taleo ID of the respondent  //
//--------------------------------------------------------------------------------------------------------------------//

function calcScores(formResponse) {
  //Declare variables
  var gritScore = 0, asqScore = 0;
  var finalGritScore = 0.0, finalAsqScore = 0.0;
  var returnObject;
  
  //----------------------------------------------//
  //  This section will calculate the Grit score  //
  //----------------------------------------------//
  
  //Loop through 12 grit assessment items (3 text fields in beginning (id, first name and last name) are skipped over) 
  for (var i = 3; i < 15; i++) {
    var itemResponse = formResponse[i];
    if (i==3 || i==6 || i==8 || i==11 || i==12 || i==14) {
      switch (itemResponse.getResponse()) {
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
    } else if (i==4 || i==5 || i==7 || i==9 || i==10 || i==13) {
      switch (itemResponse.getResponse()) {
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
  
  //Loop through items 15-38 for the CoNeg questions
  for (var i = 15; i < 39; i++) {
    var itemResponse = formResponse[i];
    if (itemResponse.getItem().getType() == 'SCALE') {
      asqScore += parseInt(itemResponse.getResponse());
    }
  }
  
  //Return final ASQ score
  finalAsqScore = asqScore / 6;
    
  //--------------------------------------------------------------------------------------------------//
  //  This section creates a JSON object to hold the scores and essay questions for sending to Taleo  //
  //--------------------------------------------------------------------------------------------------//
  
  return returnObject = JSON.stringify({ "taleoId" : formResponse[0].getResponse(), "grit" : finalGritScore.toFixed(2), "asq" : finalAsqScore.toFixed(2), "essays" : { "essay1" : escapeSpecialChars(formResponse[formResponse.length-3].getResponse()), "essay2" : escapeSpecialChars(formResponse[formResponse.length-2].getResponse()), "essay3" : escapeSpecialChars(formResponse[formResponse.length-1].getResponse()) } });
}


//-----------------------------------------------------------------------//
//  This function gets the response id from the last submitted response  //
//-----------------------------------------------------------------------//

function getFormResponseId(formId) {
  //Declare variables and get id of last form response
  var form = FormApp.openById(formId);
  var formResponses = form.getResponses();
  var lastResponse = formResponses[formResponses.length - 1];
  var lastResponseId = lastResponse.getId();
  
  //Return response id
  return lastResponseId;
}


//---------------------------------------------------------------------------------------------------------------//
//  This function replaces control characters found in the passed in string and is used for the essay responses  //
//---------------------------------------------------------------------------------------------------------------//

function escapeSpecialChars(str) {
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

function checkForDups(taleoId, formId) {
  Logger.log('Checking for previous submissions.');
  
  //Open form and grab all responses
  var form = FormApp.openById(formId);
  var formResponseSet = form.getResponses();
  
  var count = 0;
  
  //Loop through Taleo IDs to check for a match
  for(var i = 0; i < formResponseSet.length; i++) {
    var existingTaleoId = formResponseSet[i].getItemResponses()[0].getResponse();
    if(taleoId == existingTaleoId) {
      count ++;
    }
  }
  
  if(count > 1) {
    return true;
  }
  else {
    return false;
  }
}


//-----------------------------------//
//  Log into Taleo via the REST API  //
//-----------------------------------//

function taleoLogin() {
  //Build API call
  var loginURL = 'https://ch.tbe.taleo.net/CH06/ats/api/v1/login?orgCode=CITIZENSCHOOLS&userName=rorys&password=citizen1';
  var loginOptions = {
    'method' : 'post',
    'contentType' : 'application/json'
  };
  
  //Make API call
  try {
    var loginResponse = UrlFetchApp.fetch(loginURL, loginOptions);
    if (loginResponse.getResponseCode() == 200) {
      Logger.log('Logged in!');
    }
    else {
      Logger.log('Error logging in: ' + loginResponse.getContentText());
    }
  }
  catch (e) {
    Logger.log('Could not log in: ' + e);
  }
  
  //Return full authToken key
  return authToken = 'authToken=' + JSON.parse(loginResponse).response.authToken;
}


//-------------------------------------//
//  Log out of Taleo via the REST API  //
//-------------------------------------//

function taleoLogout(authToken) { 
  //Build API call
  var logoutURL = 'https://ch.tbe.taleo.net/CH06/ats/api/v1/logout';
  var logoutOptions = {
    'method' : 'post',
    'headers' : {
      'cookie' : authToken
    },
    'contentType' : 'application/json; charset=utf-8'
  };
  
  //Make API call
  try {
    var logoutResponse = UrlFetchApp.fetch(logoutURL, logoutOptions);
    if (logoutResponse.getResponseCode() == 200) {
      Logger.log('Logged out!');
    }
    else {
      Logger.log('Error logging out: ' + logoutResponse.getContentText()); 
    }
  }
  catch (e) {
    Logger.log('Could not log out: ' + e); 
  }
}


//-----------------------------------------//
//  Submit data to Taleo via the REST API  //
//-----------------------------------------//

function taleoSubmit(authToken, data) {
  //Build API call
  try {
  var submitURL = 'https://ch.tbe.taleo.net/CH06/ats/api/v1/object/candidate/' + JSON.parse(data).taleoId;
  var payload = JSON.stringify({ "candidate" : { "gritScore_txt" : JSON.parse(data).grit, "asqScore_txt" : JSON.parse(data).asq, "ntfEssay1" : JSON.parse(data).essays.essay1, "ntfEssay2" : JSON.parse(data).essays.essay2, "ntfEssay3" : JSON.parse(data).essays.essay3 }});
  var submitOptions = {
    'method' : 'put',
    'headers' : {
      'cookie' : authToken
    },
    'contentType' : 'application/json; charset=utf-8',
    'payload' : payload
  };
  }
  catch (e) {
    Logger.log('Error creating API call: ' + e); 
  }
    
  //Make API call
  try {
    var submitResponse = UrlFetchApp.fetch(submitURL, submitOptions);
    if (submitResponse.getResponseCode() == 200) {
      Logger.log('Results submitted!');
    }
    else {
      Logger.log('Error submitting results: ' + submitResponse.getContentText());
    }
  } 
  catch (e) {
    Logger.log('Could not submit results: ' + e); 
  }
}
