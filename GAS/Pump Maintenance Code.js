//// Date created 2/8/2024 Chris Dreher - 208-610-9815

function onEdit(e) {
  // Get the active sheet
  var sheet = e.source.getSheetByName("Pump Maintenance");
  
  // Check if the edited range is in column B, not in the header row, has a new value, and user is logged in
  if (e.range.getColumn() == 2 && e.range.getRow() > 1 && e.value && !e.oldValue) {
    // Get the row of the edited cell
    var row = e.range.getRow();
    
    // Get the current date and time
    var currentDate = new Date();
    
    // Format the date and time as per your requirement
    var formattedDate = Utilities.formatDate(currentDate, "GMT", "MM/dd/yyyy");
    var formattedTime = Utilities.formatDate(currentDate, "GMT", "HH:mm:ss");
    
    // Update the adjacent cells in columns D and E
    sheet.getRange(row, 4).setValue(formattedDate);
    sheet.getRange(row, 5).setValue(formattedTime);
    
    // Get user's email
    var userEmail = getUserEmail();
    
    // Update the adjacent cell in column A with the user's email
    sheet.getRange(row, 1).setValue(userEmail);
  }
}

function getUserEmail() {
    var user = Session.getActiveUser();
    if (user) {
        return user.getEmail();
    }
    return "";
}