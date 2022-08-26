//Set this to true to send emails to the replyTo address instead of the one in the sheet (for testing)
const debugMode = false;
const productName = "Commerce Cloud Partner Slack Channels";
const eventName = "Registration Request";
const sheetName = "Form Responses 1";
const actionLink = "https://docs.google.com/spreadsheets/d/1qTB5SiZGSyOx8QXJRhSzUreMM90I3lf6D2SZkfenLeU/edit?resourcekey#gid=138161112";
const slaTime = "72 hours";

//TODO: Replace the manual stuff in Slack with ~FULL AUTOMATION~!

//NOTE: approverEmailTo and approverEmailCc can be comma-delimited lists if you need multiple people involved for coverage / awareness
//const approverEmailTo = "jonathan.tucker@salesforce.com,andrew.francis@salesforce.com";
const approverEmailTo = "zafar.mohammed@salesforce.com";
const approverEmailCc = "";
const replyTo = "tzarr@salesforce.com";

// Used in a period job
const cronJobTo = "zafar.mohammed@salesforce.com";
const cronJobSubject = "Weekday digest for new Slack subscriptions";
//TODO: Update this to include team members
const cronJobCc = "";

/*
  Assumed header text values in use:
  - Timestamp  |  | Form value
  - Email Address | Form value
  - First Name | Form value
  - Last Name | Form value
  - Implementation Partner / Agency Name | Form value
  - Which channels would you like to be added to? | Form value
  - Are you already in one of our Slack channels? | Form value
  - Status | This is for our internal use and is not included in the form
  - Notes | This is for our internal use and is not included in the form
Mapping all columns by index makes debugging Apps Script easier and can be used to get prior data for duplicate checks, etc.
*/

//Column mapping for: Timestamp
const timestampHeader = "Timestamp";
const timestampColumnIndex = getHeaderIndex(sheetName, timestampHeader);
//Column mapping for: Email Address
const emailAddressHeader = "Email Address";
const emailAddressColumnIndex = getHeaderIndex(sheetName, emailAddressHeader);
//Column mapping for: First Name
const firstNameHeader = "First Name";
const firstNameColumnIndex = getHeaderIndex(sheetName, firstNameHeader);
//Column mapping for: Last Name
const lastNameHeader = "Last Name";
const lastNameColumnIndex = getHeaderIndex(sheetName, lastNameHeader);
//Column mapping for: Partner company name
const partnerCompanyNameHeader = "Implementation Partner / Agency Name";
const partnerCompanyNameColumnIndex = getHeaderIndex(sheetName, partnerCompanyNameHeader);
//Column mapping for: Which channels would you like to be added to?
const whichChannelsHeader = "Which channels would you like to be added to?";
const whichChannelsColumnIndex = getHeaderIndex(sheetName, whichChannelsHeader);
//Column mapping for: Are you already in one of our Slack channels?
const alreadyInOneHeader = "Are you already in one of our Slack channels?";
const alreadyInOneColumnIndex = getHeaderIndex(sheetName, alreadyInOneHeader);
//Column mapping for: Status
const statusHeader = "Status";
const statusColumnIndex = getHeaderIndex(sheetName, statusHeader);
//Column mapping for: Notes
const notesHeader = "Notes";
const notesColumnIndex = getHeaderIndex(sheetName, notesHeader);

// Takes a date from the sheet in a format like this and gives you whole days elapsed since that time: "3/15/2022 15:52:11"
function dateDiffDays(dateString)
{
  let lastDateEpoch = Date.parse(dateString);
  let currentDateEpoch = Date.now();
  let dateDeltaInDays = Math.floor((currentDateEpoch - lastDateEpoch)) / (1000 * 60 * 60 * 24);
  return dateDeltaInDays;
}

function onFormSubmit(e)
{
  // Get the row of submitted data using the mapped out column indexes
  let submission = {
    emailAddress: e.values[emailAddressColumnIndex]
    , firstName: e.values[firstNameColumnIndex]
    , lastName: e.values[lastNameColumnIndex]
    , partnerCompanyName: e.values[partnerCompanyNameColumnIndex]
    , whichChannels: e.values[whichChannelsColumnIndex]
    , alreadyInOne: e.values[alreadyInOneColumnIndex]
  };
  
  //You can avoid mistakes or spamming when first assigning this code to your form using the debugMode variable at the top of the code
  if(debugMode){
    submission.emailAddress = replyTo;
  }
  
  //Check for duplicates - if we have a record on file for the email address within 30 days of the first submmission it's a dup
  let alreadyRegisteredResult = isDuplicateRegistration(submission.emailAddress, submission.whichChannels);
  
  Logger.log(JSON.stringify(alreadyRegisteredResult));

  if(alreadyRegisteredResult.alreadyRegistered === true)
  {
    //Duplicate request email to the person registering
    MailApp.sendEmail({
      to: submission.emailAddress
      , replyTo: replyTo
      , subject: productName + " " + eventName + " (Duplicate)"
      , htmlBody: "Dear " + submission.firstName + " " + submission.lastName + ",<br /><br />Thanks for submitting your <b>" + productName + " " + eventName + "</b>. "
      + "Please note that this looks like a duplicate request. Your original request will be reviewed for approval as soon as possible if the request has not been actioned already.<br /><br />"
      + makeRegistrationTable(submission) + "<br />"
      + "Please allow " + slaTime + " for addition to the Slack channel to take place.<br />"
    });
    Logger.log("Sent email for duplicate request: " + productName + " " + eventName + " to <" + submission.emailAddress + ">");
    
    //Duplicate request notification to approver(s)
    MailApp.sendEmail({
      to: approverEmailTo
      , cc: approverEmailCc
      , replyTo: replyTo
      , subject: productName + " " + eventName + " (Duplicate)"
      , htmlBody: submission.firstName + " " + submission.lastName + " submitted a duplicate <b>" + productName + "</b> " + eventName + ". Further review may be required.<br /><br />"
      + makeRegistrationTable(submission) + "<br /><br />"
      + "<a href=\"" + actionLink + "\">Click here</a> to review the "+ productName + " " + eventName + ".<br />"
    });
  }
  else
  {
    // NOTE: If your form allows editing, you will get an exception for not having a recipient here when editing - so don't allow editing on your form settings! ;)
    // TODO: Figure out how to detect via code if editing of responses is  allowed and kick out an error
    
    // Only provide Slack access to those who provided channels
    if(submission !== null && submission.whichChannels !== null)
    {
      // Initial request email to the person registering
      MailApp.sendEmail({
        to: submission.emailAddress
        , replyTo: replyTo
        , subject: productName + " " + eventName + " (Initial)"
        , htmlBody: "Dear " + submission.firstName + " " + submission.lastName + ",<br /><br /> Thanks for submitting your <b>" + productName + "</b> " + eventName + ". "
        + "We now have the following data on file:<br /><br />"
        + makeRegistrationTable(submission) + "<br />"
        + "Please allow " + slaTime + " for review and processing of your request.<br />"
      });

      // TODO: Figure out how to action these with calls to the python scripts or other automation (Jason Ulrich is the contact there for Slack)

      // Commented out Initial request email to the approver or call to automation should happen here to send the Weekday Digest instead
      /*
      MailApp.sendEmail({
        to: approverEmailTo
        , cc: approverEmailCc
        , replyTo: replyTo
        , subject: "Action Required: " + productName + " " + eventName + " Review (New)"
        , htmlBody: submission.firstName + " " + submission.lastName + " submitted a new <b>" + productName + "</b> " + eventName + ". <u>Review and actioning is needed</u>:<br /><br />"
        + makeRegistrationTable(submission) + "<br /><br />"
        + "<a href=\"" + actionLink + "\">Click here</a> to review and action the request or you can respond to the daily batch job which is sent.<br />"
      });
      Logger.log("Sent email for initial request: " + productName + " " + eventName + " to approvers (To: " + approverEmailTo + " Cc: " + approverEmailCc + ")");
      */
    }
    else
    {
      MailApp.sendEmail({
        to: submission.emailAddress
        , replyTo: replyTo
        , subject: productName + " " + eventName + " (Initial)"
        , htmlBody: "Dear " + submission.fullName + ",<br /><br /> Thanks for submitting your <b>" + productName + "</b> " + eventName + ".<br /><br />" + 
        "Based on your responses we are unable to provision the access for you at this time."
      });
      Logger.log("Sent denial email for initial request: " + productName + " " + eventName + " to <" + submission.emailAddress + ">");
    }
  }
}

function getHeaderIndex(sheetName, headerText)
{
  let result = -1;
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if(sheet == null)
  {
    console.log('Sheet not located in getIndexByHeader using: ' + sheetName);
    return result;
  }

  let data = sheet.getDataRange().getValues();

  for (let i=0; i < data[0].length; i++){
    if(debugMode){
      Logger.log("Left: '" + data[0][i].toString() + "' Right: '" + headerText + "'");
    }

    if(data[0][i].toString() == headerText){
      result = i;
      break;
    }
  }

  return result;
}

// Makes an old-school HTML table of name-value pairs from the object
function makeRegistrationTable(data)
{
  Logger.log("makeRegistrationTable(" + JSON.stringify(data) + ") invoked.");
  let registrationTable = "";
  registrationTable += "<table>";
  registrationTable += "<tr><td align=\"right\"><b>" + emailAddressHeader + "</b>:</td><td>" + data.emailAddress + "</td></tr>";
  registrationTable += "<tr><td align=\"right\"><b>" + firstNameHeader + "</b>:</td><td>" + data.firstName + "</td></tr>";
  registrationTable += "<tr><td align=\"right\"><b>" + lastNameHeader + "</b>:</td><td>" + data.lastName + "</td></tr>";
  registrationTable += "<tr><td align=\"right\"><b>" + partnerCompanyNameHeader + "</b>:</td><td>" + data.partnerCompanyName + "</td></tr>";
  //TODO: Add breaks into the data.whichChannels instead of commas 
  registrationTable += "<tr><td><b>" + whichChannelsHeader + "</b></td><td> " + data.whichChannels + "</td></tr>";
  registrationTable += "<tr><td><b>" + alreadyInOneHeader + "</b></td><td>" + data.alreadyInOne + "</td></tr>";
  registrationTable += "</table>";
  return registrationTable;
}

// Detect duplicates based on composite of the Email Address and Which Channels
function isDuplicateRegistration(emailAddress, whichChannels)
{
  Logger.log("isDuplicateRegistration(" + emailAddress + ", " + whichChannels + ") invoked.");
  //The current submission does not get stopped so we need to just track it and kill the real dup looking at any matches added beyond a length of 1
  let hits = [];
  let sheet = SpreadsheetApp.getActiveSheet();
  let data  = sheet.getDataRange().getValues();
  let i = 1;
     
  // Start loop at 1 to ignore the header
  for (i = 1; i < data.length; i++)
  {
    Logger.log("Comparing: " + data[i][emailAddressColumnIndex].toString().toLowerCase() + " to: " + emailAddress.toLowerCase() + ".");
    Logger.log("Comparing: " + data[i][whichChannelsColumnIndex].toString().toLowerCase() + " to: " + whichChannels.toLowerCase() + ".");
    
    //If we have a matching email + matching whichChannels it's a duplicate
    if (data[i][emailAddressColumnIndex].toString().toLowerCase() === emailAddress.toLowerCase() && data[i][whichChannelsColumnIndex].toString().toLowerCase() === whichChannels.toLowerCase())
    {
      hits.push(data[i]);
    }
  }
  
  // Return some enriched results we can use with a boolean and reuse the data in responses
  if(hits.length > 1)
  {
    Logger.log("Located a duplicate based on partner email address: <" + emailAddress + "> and which channels of: " + whichChannels + ".");
    //Remove the newest addition using the tracking array (like it never happened)
    sheet.deleteRow(i);
    //Get the data for the original (old) submission into something we can use
    let originalRow = hits[hits.length - 2];
    let priorRegistration = {
      emailAddress: originalRow[emailAddressColumnIndex]
    , firstName: originalRow[firstNameColumnIndex]
    , lastName: originalRow[lastNameColumnIndex]
    , partnerCompanyName: originalRow[partnerCompanyNameColumnIndex]
    , whichChannels: originalRow[whichChannelsColumnIndex]
    , alreadyInOne: originalRow[alreadyInOneColumnIndex]
    };
    Logger.log("Located a duplicate based on partner email address: <" + emailAddress + "> and which channels of: " + whichChannels + ".");
    return JSON.parse("{\"alreadyRegistered\": true, \"priorData\": " + JSON.stringify(priorRegistration) +"}");
  }
  else
  {
    Logger.log("Did not find duplicate email: <" + emailAddress + "> and which channels of: " + whichChannels.toLowerCase() + ".");
    return JSON.parse("{\"alreadyRegistered\": false, \"priorData\": null}");
  }
}

function getChannelDetails(key)
{
  //TODO: This map needs to be maintained when a new channel comes online  
  channelDetailsArray = 
  [
    {
      "googleFormKey": "B2B2C (cc-b2b2c-partner-training)",
      "slackChannelId": "C03GX0894FN",      
      "owner": "Tom Zarr (tzarr@salesforce.com)",
      "isOpenToJoin": true
    },
    {
      "googleFormKey": "PWA Web Kit & Managed Runtime (cc-pwakit-mrt-training)",
      "slackChannelId": "C02E909K718",      
      "owner": "Jonathan Tucker (jonathan.tucker@salesforce.com)",
      "isOpenToJoin": true
    },
    {
      "googleFormKey": "Commerce Marketplace (Atonit) (cc-marketplace-previously-atonit-training)",
      "slackChannelId": "C03PSMLCCMC",      
      "owner": "Tom Zarr (tzarr@salesforce.com)",
      "isOpenToJoin": true
    },
    {
      "googleFormKey": "Salesforce Days '22 B2B Commerce (sfdays-commerce-b2b)",
      "slackChannelId": "C03HLG70EJH",      
      "owner": "Tom Zarr (tzarr@salesforce.com)",
      "isOpenToJoin": true
    },
    {
      "googleFormKey": "Salesforce Days '22 B2C Commerce (sfdays-commerce-b2c)",
      "slackChannelId": "C03H8SB1SUF",     
      "owner": "Jonathan Tucker (jonathan.tucker@salesforce.com)",
      "isOpenToJoin": true
    },
    {
      "googleFormKey": "Salesforce Days '22 Order Management (sfdays-commerce-oms)",
      "slackChannelId": "C03JD79PSC8",      
      "owner": "Andrew Francis (andrew.francis@salesforce.com)",
      "isOpenToJoin": true
    },
    {
      "googleFormKey": "Test (cc-test-for-script)",
      "slackChannelId": "None",      
      "owner": "Tom Zarr (tzarr@salesforce.com)",
      "isOpenToJoin": false
    }
  ];

  let result = channelDetailsArray.reduce(function(prev, curr) { return (curr.googleFormKey === key) ? curr : prev; }, null);
  
  if(result == null)
  {
    return {
      "googleFormKey": "No mapping found in channelDetailsArray",
      "slackChannelId": "No mapping found in channelDetailsArray",      
      "owner": "No mapping found in channelDetailsArray",
      "isOpenToJoin": false
    };
  }

  return result;
}

function addToChannelMap(channelMap, newValues, emailAddress)
{
  if(debugMode) {
    Logger.log("addToChannelMap(" + channelMap + "," + newValues + "," + emailAddress + ") invoked.");
  }
  
  // Populate the keys using the group values and then make arrays with emails in them for each group (key) in a single sweep
  const newValueArray = newValues.split(', ');
  let i = 0;
  
  while(i < newValueArray.length)
  {
    //Add new array with any new key
    if(!channelMap.has(newValueArray[i])){
      if(debugMode){
        Logger.log("Key not found. Adding key '" + newValueArray[i] + "' with empty array");
      }
      channelMap.set(newValueArray[i],[]);
    }

    // Append emails to the list at that key if not present
    if(channelMap.get(newValueArray[i]) != null && !channelMap.get(newValueArray[i]).includes(emailAddress)){
      if(debugMode){
        Logger.log("Adding emailAddress: '" + emailAddress + "' to array of key '" + newValueArray[i] + "'.");
      }
      channelMap.get(newValueArray[i]).push(emailAddress);
    }
    
    i++;
  }

  return channelMap;
}

function onlyUnique(value, index, self)
{
  return self.indexOf(value) === index;
}

function sendDigestEmail()
{
  //TODO: Make this send each weekday or Mon, Weds, Fri - something like that at a fixed time
  
  if(debugMode) {
    Logger.log("sendDigestEmail() invoked.");
  }
  
  let sheet = SpreadsheetApp.getActiveSheet();
  let data  = sheet.getDataRange().getValues();
  let i = 1;
     
  // Note the scan start for a null/blank status column value
  Logger.log("Scanning for a blank status column");

  // Map a dictionary of the groups with the group name as key and a list of all the emails under each group that need to be added
  const channelMap = new Map();
  
  // Start loop at 1 to ignore the header and build a map with the groups as keys
  for (i = 1; i < data.length; i++)
  {
    // Only build a map based on unactioned stuff - the rest is likely ancient history
    if (data[i][statusColumnIndex] == null || data[i][statusColumnIndex].toString() === ""){
      Logger.log("Blank status found for: " + data[i][emailAddressColumnIndex].toString().toLowerCase() + " and channels: " + data[i][whichChannelsColumnIndex].toString() + ".");
      //Add to the map
      if(data[i][whichChannelsColumnIndex] != null && data[i][whichChannelsColumnIndex].toString() !== ""){
        addToChannelMap(channelMap, data[i][whichChannelsColumnIndex].toString(), data[i][emailAddressColumnIndex].toString());
      }
    }
  }

  if(debugMode){
    // Note the scan end for a null/blank status column value
    Logger.log("Scanning for a blank status column complete.");
    Logger.log(channelMap);
  }

  // Assemble a payload broken out by group suitable for HTML email
  Logger.log("payload broken out by group suitable for HTML email.");
  let j = 0;
  let buffer = "";
  let userConversions = [];
  
  channelMap.forEach((value, key) => {
    buffer += "<hr /><b>" + key + "</b><br />";
    let details = getChannelDetails(key);
    let link = "https://salesforce-external.slack.com/archives/" + details.slackChannelId;
    let statusText = details.isOpenToJoin
      ? "<span style=\"background-color: #c1fca9\">Open to Join</span>"
      : "<span style=\"background-color: #f59393\">Closed</span>";
    buffer += "Status: <i>" + statusText + "</i><br />Channel Id: <i><a href=\"" + link + "\">" + details.slackChannelId + "</a></i><br />Owner: <i>" + details.owner + "</i><br />";
    buffer += "<b>Entries</b>:<br />";
    let k = 0;
    if(value.length == 0){
      buffer += "No new entries<br />";
    }
    else{
      while(k < value.length){
        userConversions.push("/convert_guest " + value[k]);
        buffer += value[k] + "<br />";
        k++;
      }
    }
  });

  let uniqueUserConversions = userConversions.filter(onlyUnique);
  let userConversionsString = uniqueUserConversions.join('<br />');
  let userConversionInstructions = '';

  let instructions = "<b>Instructions</b>: <a href=\"https://salesforce.quip.com/ADOOAA2yaFhF\">Click here</a>";
  let spreadsheet = "<b>Spreadsheet</b>: <a href=\"" + actionLink + "\">Click here</a>";
  
  if(userConversionsString != null && userConversionsString.length > 0){
    userConversionInstructions += "<hr /><b>Unique list of user conversion commands (single channel to multi-channel) if needed</b>:<br />";
    userConversionInstructions += userConversionsString;
  }
  
  let htmlBody = (buffer !== "")
  ? "Here is the periodic digest for external Slack channel additions. See Instructions and Spreadsheet links below:<br /><br />" + instructions + "<br />" +
  spreadsheet + "<br /><br />Summary follows...<br /><br />" + buffer + "<br />" + userConversionInstructions + "<br />"
  : "Good news! There's nothing new to add unless additional Slack channels have been opened up.";

  Logger.log("Sending email broken out by group with instructions.");
  MailApp.sendEmail({
    to: cronJobTo
    , replyTo: replyTo
    , subject: cronJobSubject
    , htmlBody: htmlBody
  });
}
