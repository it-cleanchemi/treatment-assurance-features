//Revision 8/03/2023 - with emailing stage report -- 

var TA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Treatment Assurance Reporting');
var activeCell = TA.getActiveCell(); //TA.getRange("B84");
var activeRow = activeCell.getRow();
var lastColumn = TA.getLastColumn();
var rowValues = TA.getRange(activeRow, 1, 1, lastColumn).getValues(); //.pop();
var headerRange = TA.getRange(1, 1, 1, TA.getLastColumn());
var headerValues = headerRange.getValues()[0];
const SS = SpreadsheetApp.getActiveSpreadsheet();
const WEBHOOK = SS.getSheetByName('Reference').getRange("L2").getValues();
function sendREPORT() {
    var reportSentColumnIndex = headerValues.indexOf("Report Sent") + 1;
    var referenceSheet = SS.getSheetByName('Reference');
    var lastReportRow = referenceSheet.getRange("M:M").getValues().filter(String).length;
    var REPORT = referenceSheet.getRange("M2:M" + lastReportRow).getValues();
    for (var i = 0; i < rowValues.length; i++) {
        var rowData = rowValues[i];
        var message = "-Stage Report-\n";
        var missingValues = [];
        for (var i = 0; i < REPORT.length; i++) {
            var header = REPORT[i][0]; // Access the header value from the nested array
            var value = rowData[headerValues.indexOf(header)]; // Find the corresponding value based on the header
            if (value !== null && value !== "") {
                if (value instanceof Date) {
                    var hours = value.getHours();
                    var minutes = value.getMinutes();
                    var localTime = hours.toString().padStart(2, '0') + ':' + minutes.toString().padStart(2, '0');
                    message += header + ":\t\t" + localTime + "\n";
                }
                else if (typeof value === 'number') {
                    var trimmedValue = value;
                    if (value % 1 !== 0) { // Check if the value has decimals
                        trimmedValue = value.toFixed(1); // Trim all decimals except one
                    }
                    message += header + ":\t\t" + trimmedValue + "\n";
                }
                else {
                    //   Compose the message with right-aligned header and value
                    message += header + ":\t\t" + value + "\n";
                }
            }
            else {
                missingValues.push(header); // Add the missing value header to the array
            }
        }
        var reportSent = rowData[headerValues.indexOf("Report Sent")];
        var date = new Date();
        var repDay = date.getDate();
        var repMonth = date.getMonth() + 1;
        var repYear = date.getFullYear();
        var repHours = date.getHours();
        var repMinutes = date.getMinutes();
        var repDateTime = repMonth.toString().padStart(2, '0') + "-" + repDay.toString().padStart(2, '0') + "-" + repYear.toString().padStart(2, '0') + "  " + repHours.toString().padStart(2, '0') + ':' + repMinutes.toString().padStart(2, '0');
    }
    if (missingValues.length > 0) {
        var missingValueMessage = "The following value(s) are missing:     " + "\n\n" + missingValues.join(", ") + "      Press OK to send anyway.";
        var response = Browser.msgBox("Missing Value(s)", missingValueMessage, Browser.Buttons.OK_CANCEL);
        if (response == "ok") {
            if (reportSent == "") {
                sendMessage_(WEBHOOK, message);
                TA.getRange(activeRow, reportSentColumnIndex).setValue(repDateTime);
            }
            else {
                var formattedReportSent = Utilities.formatDate(reportSent, 'America/Chicago', 'HH:mm MM-dd-yyyy');
                var response = Browser.msgBox("Attention!", "Report for this stage was posted at:  " + formattedReportSent + ".  Press OK to send it again?", Browser.Buttons.OK_CANCEL);
                // Check user response
                if (response == "ok") {
                    // Send the report
                    sendMessage_(WEBHOOK, message);
                    TA.getRange(activeRow, reportSentColumnIndex).setValue(repDateTime);
                }
            }
        }
        else {
            return;
        }
    }
    else {
        if (reportSent == "") {
            sendMessage_(WEBHOOK, message);
            TA.getRange(activeRow, reportSentColumnIndex).setValue(repDateTime);
        }
        else {
            var formattedReportSent = Utilities.formatDate(reportSent, 'America/Chicago', 'HH:mm MM-dd-yyyy');
            var response = Browser.msgBox("Attention!", "Report for this stage was posted at:  " + formattedReportSent + ".  Press OK to send it again?", Browser.Buttons.OK_CANCEL);
            // Check user response
            if (response == "ok") {
                // Send the report
                sendMessage_(WEBHOOK, message);
                TA.getRange(activeRow, reportSentColumnIndex).setValue(repDateTime);
            }
        }
    }
    var emailRecipientsRange = SS.getSheetByName('Reference').getRange("N2:N15");
    var emailRecipients = emailRecipientsRange.getValues().flat().filter(email => email !== "");
    var subject = "--Clean Chemistry Stage Report-- Well: " + rowData[5] + " Stage: " + rowData[6]; // Specify the subject of the email
    try{
     MailApp.sendEmail({
        to: emailRecipients.join(','),
        subject: subject,
        body: message,
        });
    } catch{
      
      return;
    }
}

function postShiftReport() {
    var userEmail = getUserEmail();
    var dayNight = "Day Shift";
    var ATP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ATP');
    var atpRange = ATP.getRange("A4:G" + ATP.getLastRow());
    var atpValues = atpRange.getValues();
    var atpHeaderRange = ATP.getRange(1, 1, 1, TA.getLastColumn());
    var activeRange = TA.getActiveRange();
    var lastColumn = TA.getLastColumn();
    var startRow = activeRange.getRow();
    var numRows = activeRange.getNumRows();
    var range = TA.getRange(startRow, 1, numRows, lastColumn);
    var stagesInShift = 0;
    var shiftBarrels = 0;
    var pMaxDose = 0;
    var dDACDose = 0;
    var scaleDose = 0;
    var preORP = 0;
    var postORP = 0;
    var residual = 0;
    var notes = "";
    var latestATPRaw = 0;
    var latestATPTreated = 0;
    var latestATPcomb = 0;
    var shiftReportValues = range.getValues();
    for (var i = 0; i < numRows; i++) {
        for (j = 0; j < headerValues.length; j++) {
            var currentheader = headerValues[j];
            if ((currentheader == "Stage" & shiftReportValues[i][j] > 0) || (currentheader == "1st Stage" & shiftReportValues[i][j] > 0) || (currentheader == "2nd Stage" & shiftReportValues[i][j] > 0)) {
                stagesInShift = stagesInShift + 1;
            }
            if (currentheader.includes("Total BBLs Treated per stage") || currentheader.includes("Total Treated")) {
                shiftBarrels = shiftBarrels + shiftReportValues[i][j];
            }
            if (currentheader.includes("PMAX Dose (PPM)") & shiftReportValues[i][j] > 0) {
                pMaxDose = pMaxDose + shiftReportValues[i][j];
            }
            if (currentheader.includes("DDAC Dose (PPM)") & shiftReportValues[i][j] > 0) {
                dDACDose = dDACDose + shiftReportValues[i][j];
            }
            if ((currentheader.includes("Scale") & currentheader.includes("Dose (PPM)")) & shiftReportValues[i][j] > 0) {
                scaleDose = scaleDose + shiftReportValues[i][j];
            }
            if (currentheader.includes("Pre ORP") & shiftReportValues[i][j] > 0) {
                preORP = preORP + shiftReportValues[i][j];
            }
            if (currentheader.includes("ORP - Working Tank") & shiftReportValues[i][j] > 0) {
                postORP = postORP + shiftReportValues[i][j];
            }
            if (currentheader.includes("PAA Residual") & shiftReportValues[i][j] > 0) {
                residual = residual + shiftReportValues[i][j];
            }
            if (currentheader == "Notes" & shiftReportValues[i][j] !== "") {
                notes = notes + shiftReportValues[i][j] + "\n";
            }
            var lastStageEnd = shiftReportValues[i][2];
            var lastStageEndHours = lastStageEnd.getHours();
            if (lastStageEndHours <= 8) {
                dayNight = "Night Shift";
            }
            else {
                dayNight = "Day Shift";
            }
        }
    }
    pMaxDose = pMaxDose / numRows;
    pMaxDose = pMaxDose.toFixed(1);
    if (dDACDose > 0) {
        dDACDose = dDACDose / numRows;
        dDACDose = dDACDose.toFixed(1);
    }
    else {
        dDACDose = "No DDAC used";
    }
    ;
    if (scaleDose > 0) {
        scaleDose = scaleDose / numRows;
        scaleDose = scaleDose.toFixed(1);
    }
    else {
        scaleDose = "No Scale Inhibitor used";
    }
    ;
    if (preORP > 0) {
        preORP = preORP / numRows;
        preORP = preORP.toFixed(1);
    }
    else {
        preORP = "Pre ORP is not measured";
    }
    ;
    if (postORP > 0) {
        postORP = postORP / numRows;
        postORP = postORP.toFixed(1);
    }
    else {
        postORP = "Final ORP is not measured";
    }
    ;
    if (residual > 0) {
        residual = residual / numRows;
        residual = residual.toFixed(1);
    }
    else {
        residual = "Final PAA residual is not measured";
    }
    ;
    if (notes === "") {
        var ui = SpreadsheetApp.getUi();
        ui.alert('Notes are empty', 'Please enter some highlights for your shift.', ui.ButtonSet.OK);
        return; // Exit the script
    }
    for (k = 0; k < atpValues.length; k++) {
        var comb = atpValues[k][2];
        if (comb == "") {
            continue;
        }
        else {
            latestATPcomb = atpValues[k][2];
        }
        if (atpValues[k][3] == "Raw") {
            latestATPRaw = atpValues[k][4];
        }
        if (atpValues[k][3] == "Treated") {
            latestATPTreated = atpValues[k][4];
        }
    }
    if (latestATPcomb === 0) {
        var latestATPYear = "";
        var latestATPmonth = "ATP testing is not done";
        var latestATPdate = "";
        var latestATPhours = "";
        var latestATPminutes = "";
    }
    else {
        var latestATPYear = latestATPcomb.getFullYear();
        var latestATPmonth = latestATPcomb.getMonth() + 1;
        var latestATPdate = latestATPcomb.getDate();
        var latestATPhours = latestATPcomb.getHours();
        var latestATPminutes = latestATPcomb.getMinutes();
        var formattedMinutes = Utilities.formatString('%02d', latestATPminutes);
    }
    var payload = {
        cards: [{
                header: {
                    title: "Shift Report",
                    subtitle: dayNight,
                    imageUrl: "https://fonts.gstatic.com/s/e/notoemoji/15.0/1f9a6/512.png=s60",
                    imageStyle: "IMAGE"
                },
                sections: [{
                        widgets: [{
                                textParagraph: {
                                    text: "Reported by:\t\t" + "<font color=\"#22CB7F\"> " + userEmail + "</text>" +
                                        "\n" + "<font color=\"#000000\"> Stages treated during shift:\t\t</text>" + "<font color=\"#FF0000\"> " + stagesInShift + "</text>" +
                                        "\n" + "\n" + "<font color=\"#000000\"> Volume treated (bbl):\t\t</text>" + "<font color=\"#0000FF\"> " + shiftBarrels + "</text>" +
                                        "\n" + "<font color=\"#000000\"> PeroxyMAX dose (ppm):\t\t</text>" + "<font color=\"#0000FF\"> " + pMaxDose + "</text>" +
                                        "\n" + "<font color=\"#000000\"> DDAC dose (ppm):\t\t</text>" + "<font color=\"#0000FF\"> " + dDACDose + "</text>" +
                                        "\n" + "<font color=\"#000000\"> Scale inhibitor dose (ppm):\t\t</text>" + "<font color=\"#0000FF\"> " + scaleDose + "</text>" +
                                        "\n" + "<font color=\"#000000\"> Pre-treatment ORP (mV):\t\t</text>" + "<font color=\"#0000FF\"> " + preORP + "</text>" +
                                        "\n" + "<font color=\"#000000\"> Post-treatment ORP (mV):\t\t</text>" + "<font color=\"#0000FF\"> " + postORP + "</text>" +
                                        "\n" + "<font color=\"#000000\"> Post-treatment PAA residual (ppm):\t\t</text>" + "<font color=\"#0000FF\"> " + residual + "</text>" +
                                        "\n\n" + "<font color=\"#000000\"> Latest ATP (pg/mL) is measured on " + latestATPmonth + "-" + latestATPdate + "-" + latestATPYear + " at " + latestATPhours + ":" + formattedMinutes + " " +
                                        "\n" + "Raw:\t\t </text>" + "<font color=\"#0000FF\"> " + latestATPRaw + "</text>" +
                                        "\n" + "<font color=\"#000000\"> Treated:\t\t</text>" + "<font color=\"#0000FF\"> " + latestATPTreated + "</text>" +
                                        "\n" + "<font color=\"#000000\"> Notes:\n</text>" + "<font color=\"#0000FF\"> " + notes + "</text>"
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
    var response = UrlFetchApp.fetch(url, options);
}
function sendMessage_(webhook, message) {
    // Sends the message text to the given webhook URL
    const payload = JSON.stringify({ text: message });
    const options = {
        method: 'POST',
        contentType: 'application/json',
        payload: payload,
    };
    UrlFetchApp.fetch(webhook, options);
}
function getUserEmail() {
    var user = Session.getActiveUser();
    if (user) {
        return user.getEmail();
    }
    return "";
}
