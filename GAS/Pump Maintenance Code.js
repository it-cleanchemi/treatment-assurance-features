/*

Version 2/23/2024

Edited by v.martysevich@cleanchemi.com

The script will:
1 require to fill up the change data in "Pump Maintenance" tab,
1a populate time and date on edit
2 display the waring if the data incomplete,
3. post note to the job chat
4. Populate record sheet
5. erase hose from inventory
6 clean the "Pump Maintenance" tab,



 */

//// Date created 2/8/2024 Chris Dreher - 208-610-9815
const SSS = SpreadsheetApp.getActiveSpreadsheet();
var pmSheet = SSS.getSheetByName('Pump Maintenance');
var recordSheet = SSS.getSheetByName('Pump Maintenance Records');
var referenceSheet = SSS.getSheetByName('Reference');



function onEdit(e) {
  // Call the helper function with the real event object
  processEdit(e);
}

function processEdit(e) {
  // Your original logic here, slightly modified to use a passed event object
  var editedSheet = e.range.getSheet().getName();
  var targetSheet = pmSheet.getName();

  if (editedSheet == targetSheet){ // Assuming pmSheet is a sheet name
    if (e.range.getColumn() == 1 && e.range.getRow() == 2 && e.value && !e.oldValue) {
      var currentDate = new Date();
      var formattedDate = Utilities.formatDate(currentDate, "CST", "MM/dd/yyyy");
      var formattedTime = Utilities.formatDate(currentDate, "CST", "HH:mm:ss");
      
      pmSheet.getRange(2, 3).setValue(formattedDate);
      pmSheet.getRange(2, 4).setValue(formattedTime);
    }
  }
}


function hoseChange() {
  //receive hose change info
  var hoseChangeDataRange = pmSheet.getRange('A2:H2');
  var hoseChangeData = hoseChangeDataRange.getValues();
  var hoseID = hoseChangeData[0][0];
  var hoseType = hoseChangeData[0][1];
  var changeDate = hoseChangeData[0][2];
  var changeTime = hoseChangeData[0][3];
  var changeHours = hoseChangeData[0][4];
  var changePump = hoseChangeData[0][5];
  var changeReason = hoseChangeData[0][6];
  var changeNotes = hoseChangeData[0][7];
  
  // Check for missing required data (all except changeNotes)
  if (!hoseID || !hoseType || !changeDate || !changeTime || !changeHours || !changePump || !changeReason) {
    // Display an alert to the user
    var ui = SpreadsheetApp.getUi(); // Get the UI of the spreadsheet
    ui.alert('Missing Data', 'Please ensure all hose change data is filled out.', ui.ButtonSet.OK);
    return; // Exit the function early
  }
  
  // Get user's email
  var userEmail = getUserEmail();
  var recordRow = recordSheet.getLastRow()+1
  recordSheet.getRange(recordRow, 1).setValue(userEmail);
  

  //Populating hose change log

  var numberOfColumns = hoseChangeData[0].length; 
  var recordRange = recordSheet.getRange(recordRow, 2, 1, numberOfColumns);
  recordRange.setValues([hoseChangeData[0]]);
  // Cleaning inventory

  // Find hoseID in referenceSheet and clear values in columns O and P for that row
  var referenceDataRange = referenceSheet.getRange('O:O');
  var referenceData = referenceDataRange.getValues(); // This gets a 2D array of all values in column O
  
  for (var i = 0; i < referenceData.length; i++) {
    if (referenceData[i][0] == hoseID) {
      // Found the row with hoseID, now clear values in columns O and P for this row
      // Rows and columns are 1-indexed in getRange, so adjust accordingly
      referenceSheet.getRange(i + 1, 15, 1, 2).clearContent(); // 15 is column O, 2 is the number of columns to clear (O and P)
      break; // Exit the loop once the match is found and cleared
    }
  }
  // Cleaning Hose change form
  pmSheet.getRange('A2').clearContent();
  pmSheet.getRange('C2:H2').clearContent();


  //Send a report

 var cardMessage = {
    "cards": [{
      "header": {
        "title": "Hose Change Alert",
        "subtitle": "Pump Maintenance System"
      },
      "sections": [{
        "widgets": [{
          "keyValue": {
            "topLabel": "Pump",
            "content": changePump
          }
        }, {
          "keyValue": {
            "topLabel": "Hose Type",
            "content": hoseType
          }
        }, {
          "keyValue": {
            "topLabel": "Reason",
            "content": changeReason
          }
        }, 
        // Adding changeHours to the card message
        {
          "keyValue": {
            "topLabel": "Hours",
            "content": changeHours.toString() // Ensure it's a string; necessary if changeHours is a number
          }
        },
        {
          "textParagraph": {
            "text": "Notes: " + (changeNotes || "N/A")
          }
        }]
      }]
    }]
  };

  // Sending the card message to WEBHOOK2
  var webhookUrl = WEBHOOK[0][0]; // Assuming WEBHOOK2 is obtained via getValues() and is in the first cell
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(cardMessage)
  };

  try {
    UrlFetchApp.fetch(webhookUrl, options);
  } catch (e) {
    Logger.log("Error sending card to webhook: " + e.toString());
    // Optionally, handle this error (e.g., display an alert to the user)
  }


}




function getUserEmail() {
    var user = Session.getActiveUser();
    if (user) {
        return user.getEmail();
    }
    return "";
}