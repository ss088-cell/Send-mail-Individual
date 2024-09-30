// Hardcoded Google Sheet ID for the recipient information (replace with your actual sheet ID)
const RECIPIENT_SHEET_ID = 'your-hardcoded-sheet-id-here';  // Replace with actual sheet ID

// Fixed URL for HPS Security Dashboard
const SECURITY_DASHBOARD_URL = 'https://datastudio.google.com/u/0/reporting/your-dashboard-link-here';

// Function to extract appName from the Google Sheet name
function extractAppName(sheetUrl) {
  try {
    const sheet = SpreadsheetApp.openByUrl(sheetUrl);  // Use openByUrl to handle the full URL
    const sheetName = sheet.getName();

    // Split based on '-' and extract appName (third part)
    const parts = sheetName.split('-');
    if (parts.length >= 6 && parts[0] === "Macroscope Scan") {
      return parts[2]; // Return appName (third part)
    }

    return null;
  } catch (error) {
    Logger.log('Error extracting app name: ' + error.message);
    return null;
  }
}

// Function to fetch email details based on appName (using hardcoded sheet ID)
function fetchEmailDetails(sheetUrl) {
  const appName = extractAppName(sheetUrl);
  if (!appName) {
    return null;
  }

  // Open the specific Google Sheet using the hardcoded sheet ID for recipient data
  const recipientSpreadsheet = SpreadsheetApp.openById(RECIPIENT_SHEET_ID);
  const recipientSheet = recipientSpreadsheet.getSheetByName('Recipients'); // Sheet with email data
  const data = recipientSheet.getDataRange().getValues();

  let emailDetails = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === appName) {
      // Construct the email body with the two links
      const reportLink = `<a href="${sheetUrl}">${appName} Report</a>`;
      const emailBody = `${data[i][4].replace('{link}', reportLink)}
        \n\nRefer to LookerStudio Security Dashboard: 
        <a href="${SECURITY_DASHBOARD_URL}">HPS Security Dashboard</a>`;

      emailDetails = {
        to: data[i][1],
        cc: data[i][2],
        subject: data[i][3],
        body: emailBody
      };
      break;
    }
  }

  return emailDetails;
}

// Function to send the report email
function sendReportEmail(sheetUrl, emailDetails) {
  if (!emailDetails) {
    return false;
  }

  MailApp.sendEmail({
    to: emailDetails.to,
    cc: emailDetails.cc,
    subject: emailDetails.subject,
    htmlBody: emailDetails.body
  });

  return true;
}
