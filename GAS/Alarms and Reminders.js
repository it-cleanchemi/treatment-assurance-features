var refS = SS.getSheetByName("Reference");
var triggesRange = refS.getRange("Q2:T"+refS.getLastRow());
var triggersAll = triggesRange.getValues();
var triggers = triggersAll.filter(function(row) {
    // Check if the row has at least one non-empty element
    return row.some(function(cell) {
      return cell !== ""; // Adjust this condition if necessary to match your criteria for "empty"
    });
});

function triggerFunction(){
  
}

function remindInventory(){
  const inventorySheet = SS.getSheetByName("Active Inventory");
  const now = new Date();
  var inventoryCombinedRange = inventorySheet.getRange("A2:A500");
  var inventoryCombined = inventoryCombinedRange.getValues();
  var filteredInventory = inventoryCombined.flat().filter(element => element !== "");
  if(filteredInventory.length > 0) {
    var dates = filteredInventory.map(element => new Date(element));
    var maxDate = new Date(Math.max.apply(null, dates));
    var lastDate = new Date(filteredInventory[filteredInventory.length - 1]);
  }
  // Calculate the differences
    var diffToMaxDate = now - maxDate; // Difference to maxDate in milliseconds
    var diffToLastDate = now - lastDate; // Difference to lastDate in milliseconds

    // Convert milliseconds to more meaningful units, e.g., days
    var diffToMaxDateHours = diffToMaxDate / (1000 * 60 * 60);
    var diffToLastDateHours = diffToLastDate / (1000 * 60 * 60);

    var lastDateString = lastDate.toLocaleString();
  
  
  //Sending reminder
  if (diffToLastDateHours>15){
    var payload = {
        cards: [{
          header: {
            title: "<b> Inventory Reminder </b>",
            subtitle: "Urgent!",
            imageUrl: "https://fonts.gstatic.com/s/e/notoemoji/15.0/1f4dd/512.png=s32",
            imageStyle: "IMAGE"
          },
          sections: [{
            widgets: [{
              textParagraph: {
                text: "<font color=\"#FF0000\"> <b> Looks like chemical inventory was not done for this shift. Please update numbers in TA </b>"+ 
                "\n Your prompt response is very valuable!"+
                "\n <font color=\"#0000FF\">Last inventory is done: "+ lastDateString
              },
            }],
          }],
        }]
    }; 
    
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    };
    
    var url = WEBHOOK;
    UrlFetchApp.fetch(url, options);
  }
}


function requestORP() {

 var payload = {
    cards: [{
      header: {
        title: "<b> ORP Callibration Verification Request</b>",
        subtitle: "Urgent!",
        imageUrl: "https://fonts.gstatic.com/s/e/notoemoji/15.0/1f9ea/512.png=s32",
        imageStyle: "IMAGE"
      },
      sections: [{
        widgets: [{
          textParagraph: {
            text: "<font color=\"#FF0000\"> <b> Take a picture of the ORP probe in the standard fluid with the numbers on its screen and post it to this chat.</b>"+ 
            "\n Your prompt response is very valuable!"
          },
        }],
      }],
    }]
  }; 
  
  var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
    };
    var url = WEBHOOK;
    UrlFetchApp.fetch(url, options);
 
}

function reminderATP() {
 
 var payload = {
    cards: [{
      header: {
        title: "ATP test reminder",
        subtitle: "",
        imageUrl: "https://fonts.gstatic.com/s/e/notoemoji/15.0/2623_fe0f/512.png=s60",
        imageStyle: "IMAGE"
      },
      sections: [{
        widgets: [{
          textParagraph: {
            text: "<font color=\"#0000FF\"> <b> ATP tests are requred once per day. Please make sure you do it before the end of the shift.</b>"+ 
            "\n Remenber! Accurate reporting is very important!"
          },
        }],
      }],
    }]
  }; 
  
  var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
    };
    var url = WEBHOOK;
    UrlFetchApp.fetch(url, options);
  

  Logger.log("1");
}

function remindersConducivity() {
 var payload = {
    cards: [{
      header: {
        title: "<b> Conductivity Verification Request</b>",
        subtitle: "Urgent!",
        imageUrl: "https://fonts.gstatic.com/s/e/notoemoji/15.0/2697_fe0f/512.png=s50",
        imageStyle: "IMAGE"
      },
      sections: [{
        widgets: [{
          textParagraph: {
            text: "<font color=\"#FF0000\"> <b> Take a picture of the Conductivity probe in the calibration solution (111 mcS/cm) with the numbers on its screen and post it to this chat.</b>"+ 
            "\n Your prompt response is very valuable!"
          },
        }],
      }],
    }]
  }; 
  
  var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
    };
    var url = WEBHOOK;
    UrlFetchApp.fetch(url, options);
}

function remindersCondProbeCleaning() {
 var payload = {
    cards: [{
      header: {
        title: "Time to clean Inline probe!",
        subtitle: "",
        imageUrl: "https://fonts.gstatic.com/s/e/notoemoji/15.0/1f6e0_fe0f/512.png=s30",
        imageStyle: "IMAGE"
      },
      sections: [{
        widgets: [{
          textParagraph: {
            text: "<font color=\"#0000FF\"> <b> Inline probe needs to be cleaned ones per week. Please make sure you do it before the end of the shift.</b>"+ 
            "\n Remenber! Refer to manual or call your supervisor if you are not sure how to do it!"
          },
        }],
      }],
    }]
  }; 
  
  
  var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
    };
    var url = WEBHOOK;
    UrlFetchApp.fetch(url, options);
}