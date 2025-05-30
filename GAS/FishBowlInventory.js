/**
 Version 11/14/2024 - Centralized Database Access + retreaving Script ID function

Edited by v.martysevich@cleanchemi.com

The script will:
1 Pull data from Fishbowl database 
2 compare differences of inventory with current inventory in TA
3 Update comparison tabla in the TA
4 Communicate results via email and job chat
 * 
 * 
 * **/

 
const dashboardID = '1WWm9KFEU2ROG60TeorZxD3f6-C9FdP4VYR8g2udIpug';
const mappingSheet = 'Chemical Mapping';
const dashBoard = SpreadsheetApp.openById(dashboardID).getSheetByName(mappingSheet);
const address = dashBoard.getRange('D2').getValue();
const port = dashBoard.getRange('D3').getValue();
const dbName = dashBoard.getRange('D4').getValue();
const username = dashBoard.getRange('D5').getValue();
const password = dashBoard.getRange('D6').getValue();
const emailRecipient = dashBoard.getRange('D7').getValue();

function updateConsumption() {
    const sheet = SS.getSheetByName("Fishbowl Inventory");
    if (!sheet) {
        sheet = SS.insertSheet("Fishbowl Inventory");
    }
    var sheetName = SS.getName();
    var jobCode = sheetName.substring(0, 5);
    var jobName = extractJobDetails(sheetName);

    // Fetch data from the database
    var dbData = fetchCurrentInventory(jobCode);
    if (dbData===-1){
        Browser.msgBox("Logistics server is down, try to sync inventory later.");
        return;
    }
    // Compare inventories and prepare comparison data
    var comparisonData = compareInventories(dbData,jobCode);
    
    if (comparisonData.rows.length > 1) {
        sheet.clearContents();
        sheet.getRange(1, 1, comparisonData.rows.length, comparisonData.rows[0].length).setValues(comparisonData.rows);
    } else {
        Logger.log("No data found.");
    }
    // Generate HTML content for email
    var htmlOutput = HtmlService.createTemplateFromFile('InventoryUpdate');
    htmlOutput.tableRows = comparisonData.tableRows;
    htmlOutput.jobName = comparisonData.jobName;
    var htmlContent = htmlOutput.evaluate().getContent();

    // Send email
    
    var emailSubject = jobName +" Consumption Report";
    var emailBody = `Dear Team,<br><br>Please find below the inventory differences report for ${jobName}:<br><br>` + htmlContent + `<br><br>Best Regards,<br>Your Automated System`;

    //MailApp.sendEmail({
    //    to: emailRecipient,
    //    subject: emailSubject,
    //    htmlBody: emailBody
    //});

    // Generate summary from comparisonData
    var summary = generateSummary(comparisonData.rows);

    // Post to Google Chat
    postToGoogleChat(summary, jobName);

    Logger.log("Inventory consumption differences emailed and posted to Google Chat.");
}

function fetchCurrentInventory(jobCode) {
  
    try {
        var url = 'jdbc:mysql://' + address + ':' + port + '/' + dbName;
        var conn = Jdbc.getConnection(url, username, password);
        var stmt = conn.createStatement();

        var query = `
            WITH MaxRecord AS (
                SELECT
                    info,
                    MAX(recordid) AS max_recordid
                FROM
                    trackinginfo
                GROUP BY
                    info
            )
            SELECT
                tt.info AS LotInfo,
                p.description AS PartDescription,
                l.name AS LocationName,
                SUM(t.qty) AS TotalQuantity,
                MAX(t.dateCreated) AS dateCreated,
                t.dateLastModified AS lastCount,
                t.id AS id,
                ti.qty AS OriginalQuantity
            FROM 
                tag t
            JOIN 
                location l ON t.locationId = l.id
            JOIN 
                part p ON t.partId = p.id
            LEFT JOIN 
                trackingtext tt ON t.id = tt.tagId
            LEFT JOIN
                MaxRecord mr ON tt.info = mr.info
            LEFT JOIN
                trackinginfo ti ON tt.info = ti.info AND ti.recordid = mr.max_recordid
            WHERE
                l.name = '${jobCode}'
            GROUP BY 
                tt.info, p.description, l.name, t.id, ti.qty
            ORDER BY 
                p.description, l.name;
        `;

        var rs = stmt.executeQuery(query);
        var data = [];
        var columnCount = rs.getMetaData().getColumnCount();
        var columnNames = [];

        for (var j = 1; j <= columnCount; j++) {
            columnNames.push(rs.getMetaData().getColumnName(j));
        }
        data.push(columnNames);

        while (rs.next()) {
            var rowData = [];
            for (var k = 1; k <= columnCount; k++) {
                rowData.push(rs.getString(k));
            }
            data.push(rowData);
        }

        rs.close();
        stmt.close();
        conn.close();

        var filteredData = data.filter(function(row) {
            return row[0] !== null; // Check if the first column (index 0) is not null
        });
        
        return filteredData;
    }catch(e){
        Logger.log("Error connecting to the database or executing the query: " + e.message);
        var eReturn = -1
        return eReturn;
    } finally {
        if (conn) {
            try {
                conn.close();
            } catch (e) {
                Logger.log("Error closing the connection: " + e.message);
            }
        }
    }
}

function compareInventories(dbData, jobCode) {
    var activeInventorySheet = SS.getSheetByName("Active Inventory");
    var activeInventoryRange = activeInventorySheet.getRange('A2:J');
    var activeInventoryDataUnfiltered = activeInventoryRange.getValues();

    // Filter the data to keep only rows where the 6th column is not an empty string
    var activeInventoryData = activeInventoryDataUnfiltered.filter(function(row) {
        return row[5] !== "";
    });

    // Create a map for the latest records in active inventory based on tote identifier
    var latestActiveInventory = {};
    var firstSeenMap = {};
    var lastSeenMap = {};
    var toteDeliveryMap = {};

    activeInventoryData.forEach(function(row) {
        var toteIdentifier = row[5];
        var date = new Date(row[0]); // Assuming the date is in the first column
        var originalQuantity = row[9];

        // Update firstSeenMap only if row[2] does not contain "Delivery"
        if (!row[2].includes("Delivery")) {
            if (!firstSeenMap[toteIdentifier]) {
                firstSeenMap[toteIdentifier] = { date: date, originalQuantity: originalQuantity };
            } else if (date < firstSeenMap[toteIdentifier].date) {
                firstSeenMap[toteIdentifier] = { date: date, originalQuantity: originalQuantity };
            }

            // Update lastSeenMap only if row[2] does not contain "Delivery"
            lastSeenMap[toteIdentifier] = date;
        }else {
            // Update toteDeliveryMap for deliveries
            if (!toteDeliveryMap[toteIdentifier]) {
                toteDeliveryMap[toteIdentifier] = { date: date, originalQuantity: originalQuantity };
            }
        }

        if (!latestActiveInventory[toteIdentifier]) {
            latestActiveInventory[toteIdentifier] = row;
        } else {
            var currentLatestDate = new Date(latestActiveInventory[toteIdentifier][0]);
            if (date > currentLatestDate) {
                latestActiveInventory[toteIdentifier] = row;
            }
        }
    });

    
    
    
    // Prepare the comparison data array
    var comparisonData = [];
    var tableRows = "";
    
    var headers = ["Tote Number", "Chem Type", "Job", "Current Fishbowl Gallons", "Date Moved to Site", "Date Last Modified", "Tag ID", "Original Gallons", "Current Gallons (On Site)", "Gallons Difference", "Active & Unused Totes", "First Reported Date in use", "Last Reported Date in use"];
    comparisonData.push(headers);
    tableRows += "<tr>";
    headers.forEach(function(header) {
        tableRows += `<th>${header}</th>`;
    });
    tableRows += "</tr>";

    // Create a set of tote identifiers from dbData for quick lookup
    var dbToteIdentifiers = new Set(dbData.slice(1).map(row => row[0]));

    // Loop through the Fishbowl inventory and compare with the active inventory
    dbData.slice(1).forEach(function(row) {
        var toteIdentifier = row[0];
        var prefix = toteIdentifier.split('-')[0];  // Extract prefix from tote identifier
        var chemicalName = getChemicalName(prefix);  // Get chemical name from prefix
        var fBQuantity = parseFloat(row[3]);
        var tId = row[6];
        var originalQuantity = parseFloat(row[7]);
        var activeInventoryQuantity = latestActiveInventory[toteIdentifier] ? parseFloat(latestActiveInventory[toteIdentifier][9]) : "";
        var difference = activeInventoryQuantity !== "" ? activeInventoryQuantity - fBQuantity : "";
        var activeTote = (activeInventoryQuantity !== 0) ? toteIdentifier : "";
        var firstSeen = (firstSeenMap[toteIdentifier] && !latestActiveInventory[toteIdentifier][2].includes("Delivery")) ? formatDateTime(firstSeenMap[toteIdentifier].date) : '';
        var lastSeen = (lastSeenMap[toteIdentifier] && !latestActiveInventory[toteIdentifier][2].includes("Delivery")) ? formatDateTime(lastSeenMap[toteIdentifier]) : '';

        var rowData = [toteIdentifier, chemicalName, row[2], fBQuantity, row[4], row[5], tId, originalQuantity, activeInventoryQuantity, difference, activeTote, firstSeen, lastSeen];
        comparisonData.push(rowData);

        tableRows += "<tr>";
        rowData.forEach(function(cell, index) {
            if (index === 0) {
                tableRows += `<td style="color: blue;">${cell}</td>`;
            } else {
                tableRows += `<td>${cell}</td>`;
            }
        });
        tableRows += "</tr>";
    });

    // Check for totes in active inventory that are not in dbData and add them as empty totes
    Object.keys(latestActiveInventory).forEach(toteIdentifier => {
        if (!dbToteIdentifiers.has(toteIdentifier)) {
            var activeRow = latestActiveInventory[toteIdentifier];
            var prefix = toteIdentifier.split('-')[0];  // Extract prefix from tote identifier
            var chemicalName = getChemicalName(prefix);  // Get chemical name from prefix
            var job = jobCode;
            var firstSeen = firstSeenMap[toteIdentifier] ? formatDateTime(firstSeenMap[toteIdentifier].date) : '';
            var activeInventoryQuantity = latestActiveInventory[toteIdentifier] ? parseFloat(latestActiveInventory[toteIdentifier][9]) : "";
            var lastSeen = lastSeenMap[toteIdentifier] ? formatDateTime(lastSeenMap[toteIdentifier]) : '';
            var originalQuantity = toteDeliveryMap[toteIdentifier] ? toteDeliveryMap[toteIdentifier].originalQuantity : 0;
            var toteNameForList = "";
            var finalReport = "";
            
            if(activeInventoryQuantity===0){
              finalReport = lastSeen
            }
            if(activeInventoryQuantity != 0){toteNameForList = toteIdentifier};

            var rowData = [toteIdentifier, chemicalName, job, 0, firstSeen, lastSeen, "", originalQuantity, activeInventoryQuantity, activeInventoryQuantity, toteNameForList, firstSeen, finalReport];
            comparisonData.push(rowData);

            tableRows += "<tr>";
            rowData.forEach(function(cell, index) {
                if (index === 0) {
                    tableRows += `<td style="color: blue;">${cell}</td>`;
                } else {
                    tableRows += `<td>${cell}</td>`;
                }
            });
            tableRows += "</tr>";
        }
    });

    // Add new totes to the "Active Inventory" sheet
    var newActiveInventoryRows = [];
    dbData.slice(1).forEach(function(row) {
        var toteIdentifier = row[0];
        if (!latestActiveInventory[toteIdentifier]) {
            var chemPrefix = toteIdentifier.split('-')[0];
            var chemicalName = getChemicalName(chemPrefix)  // Extract chemical name from tote identifier
            var currentTime = new Date();
            var ampm = currentTime.getHours() >= 12 ? "Delivery PM" : "Delivery AM";
            var time = currentTime.getHours().toString().padStart(2, '0') + ":" + currentTime.getMinutes().toString().padStart(2, '0');
            var date = (currentTime.getMonth() + 1).toString().padStart(2, '0') + '/' + 
                       currentTime.getDate().toString().padStart(2, '0') + '/' + 
                       currentTime.getFullYear();
            newActiveInventoryRows.push(["", date, ampm, time, "Auto", toteIdentifier, "", "", chemicalName, row[3]]);
        }
    });

    if (newActiveInventoryRows.length > 0) {
        // Find the last row with data in column A
        var lastRow = activeInventorySheet.getRange('A:A').getValues().filter(String).length;
        
        // Add the new rows while keeping the existing formulas in columns A and B intact
        newActiveInventoryRows.forEach((newRow, index) => {
            var newRowNumber = lastRow + index + 1;
            activeInventorySheet.getRange(newRowNumber, 2).setValue(newRow[1]); // Column B
            activeInventorySheet.getRange(newRowNumber, 3).setValue(newRow[2]); // Column C
            activeInventorySheet.getRange(newRowNumber, 4).setValue(newRow[3]); // Column D
            activeInventorySheet.getRange(newRowNumber, 5).setValue(newRow[4]); // Column E
            activeInventorySheet.getRange(newRowNumber, 6).setValue(newRow[5]); // Column F
            activeInventorySheet.getRange(newRowNumber, 9).setValue(newRow[8]); // Column I
            activeInventorySheet.getRange(newRowNumber, 10).setValue(newRow[9]); // Column J
        });
    }

    return { tableRows: tableRows, rows: comparisonData };
}


function extractJobDetails(sheetName) {
    var match = sheetName.match(/^([A-Z]{2}\d{3}) Treatment Assurance Data Collection - (.+?) \((.+?)\)$/);
    if (match) {
        var code = match[1];
        var customerPadName = match[2];
        var customerName = customerPadName.split(' ')[0];
        var padName = customerPadName.split(' ').slice(1).join(' ');
        var jobName = code + " " + customerName + " " + padName;
        return jobName;
    }
    return sheetName; // Fallback in case of a non-matching format
}

function generateSummary(comparisonData) {
    // Initialize the summary object
    var chemicalSummary = {};

    // Loop through the comparison data to calculate the summary
    comparisonData.slice(1).forEach(function(row) {
        var chemicalName = row[1];
        var currentGallons = row[8] !== "" ? parseFloat(row[8]) : null;
        var fBQuantity = parseFloat(row[3]) || 0;
        var originalQuantity = parseFloat(row[7]) || 0;

        if (!chemicalSummary[chemicalName]) {
            chemicalSummary[chemicalName] = {
                fullTotes: 0,
                partialTotes: 0,
                emptyTotes: 0,
                totalGallons: 0
            };
        }

        if (currentGallons === 0) {
            // Empty totes
            chemicalSummary[chemicalName].emptyTotes++;
        } else if (currentGallons >= originalQuantity) {
            // Full totes
            chemicalSummary[chemicalName].fullTotes++;
            chemicalSummary[chemicalName].totalGallons += currentGallons;
        } else {
            // Partial totes
            chemicalSummary[chemicalName].partialTotes++;
            chemicalSummary[chemicalName].totalGallons += currentGallons;
        }
    });

    // Ensure totalGallons is always a number
    Object.keys(chemicalSummary).forEach(function(chemicalName) {
        if (isNaN(chemicalSummary[chemicalName].totalGallons)) {
            chemicalSummary[chemicalName].totalGallons = 0;
        }
    });

    return chemicalSummary;
}
function postToGoogleChat(chemicalSummary, jobName) {
  var chatWebhookUrl = WEBHOOK; //"https://chat.googleapis.com/v1/spaces/AAAApEyy8XY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=cc0XVeeTaP9QLBbrDecNG76cqBxHppp_SAROy6MTAvg";
 var emptyTotesCount = 0;

    // Build the summary text
    var text = `Location Inventory Summary for ${jobName}:\n\n`;
    Object.keys(chemicalSummary).forEach(function(chemicalName) {
        var summary = chemicalSummary[chemicalName];
        if (chemicalName === "Empty") {
            emptyTotesCount = summary.emptyTotes;
            return;
        }

        text += `<font color=\"#0000FF\"><b>${chemicalName}</b></font>\n`;
        text += `New Totes: <b>${summary.fullTotes}</b>\n`;
        text += `Partial Totes: <b>${summary.partialTotes}</b>\n`;
        if (summary.emptyTotes > 0) {
            text += `Empty Totes: <b>${summary.emptyTotes}</b>\n`;
        }
        text += `Total Gallons: <b>${summary.totalGallons.toFixed(0)}</b>\n\n`;
    });

    if (emptyTotesCount > 0) {
        text += `<b> Empty totes on location: <font color=\"#008000\"> ${emptyTotesCount}</b></font>\n\n`;
    }

    var card = {
        "cards": [
            {
                "header": {
                    "title": "Location Inventory Summary",
                    "subtitle": jobName,
                    "imageUrl": "https://fonts.gstatic.com/s/e/notoemoji/15.1/1f4dd/512.png=s64"
                },
                "sections": [
                    {
                        "widgets": [
                            {
                                "textParagraph": {
                                    "text": text
                                }
                            }
                        ]
                    }
                ]
            }
        ]
    };

    var options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(card)
    };

    UrlFetchApp.fetch(chatWebhookUrl, options);
}
// Helper function to format date and time as mm/dd/yyyy hh:mm:ss
function formatDateTime(date) {
        var mm = (date.getMonth() + 1).toString().padStart(2, '0');
        var dd = date.getDate().toString().padStart(2, '0');
        var yyyy = date.getFullYear();
        var hh = date.getHours().toString().padStart(2, '0');
        var mi = date.getMinutes().toString().padStart(2, '0');
        var ss = date.getSeconds().toString().padStart(2, '0');
        return `${mm}/${dd}/${yyyy} ${hh}:${mi}:${ss}`;
}

function getChemicalName(prefix) {
  if (!prefix) {prefix="CA50"}; //Debugging
      
  // Get the range of data (assuming data starts at A1 and has two columns)
  var dataRange = dashBoard.getRange(2, 1, dashBoard.getLastRow(), 2);
  var data = dataRange.getValues();

  // Search for the prefix in the sheet
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === prefix) {
      return data[i][1];  // Return the chemical name if the prefix is found
    }
  }

  // If no match is found, return the prefix as a fallback
  return prefix;
}

