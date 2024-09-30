// Hardcoded Google Sheet ID for the recipient information (replace with your actual sheet ID)
const RECIPIENT_SHEET_ID = 'your-hardcoded-sheet-id-here';  // Replace this with the actual sheet ID

// Function to extract appName from the Google Sheet name
function extractAppName(sheetUrl) {
  const sheet = SpreadsheetApp.openByUrl(sheetUrl);
  const sheetName = sheet.getName();

  // Check if the sheet name follows the format "Macroscope Scan - Teamname - appName"
  const parts = sheetName.split(' - ');  // Split based on ' - '

  if (parts.length === 3 && parts[0] === "Macroscope Scan") {
    return parts[2]; // Return appName (the third part)
  }

  return null; // Return null if the sheet name format doesn't match
}

// Function to fetch email details based on appName (using hardcoded sheet ID)
function fetchEmailDetails(sheetUrl) {
  const appName = extractAppName(sheetUrl);

  if (!appName) {
    return null; // Return null if appName is not found
  }

  // Open the specific Google Sheet using the hardcoded sheet ID
  const recipientSpreadsheet = SpreadsheetApp.openById(RECIPIENT_SHEET_ID);
  const recipientSheet = recipientSpreadsheet.getSheetByName('Recipients'); // Assuming the sheet is named 'Recipients'
  
  const data = recipientSheet.getDataRange().getValues();
  let emailDetails = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === appName) { // Assuming the first column contains app names
      emailDetails = {
        to: data[i][1],
        cc: data[i][2],
        subject: data[i][3],
        body: data[i][4].replace('{link}', sheetUrl) // Insert the sheet URL in the body
      };
      break;
    }
  }
  
  return emailDetails;
}

// Function to send the report email
function sendReportEmail(sheetUrl) {
  const emailDetails = fetchEmailDetails(sheetUrl);

  if (!emailDetails) {
    return false; // Return false if no email details are found
  }

  MailApp.sendEmail({
    to: emailDetails.to,
    cc: emailDetails.cc,
    subject: emailDetails.subject,
    body: emailDetails.body
  });

  return true;
}
