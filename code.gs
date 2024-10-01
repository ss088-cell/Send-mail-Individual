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
            // Construct the hardcoded email body
            const currentYear = new Date().getFullYear();
            const quarter = Math.ceil((new Date().getMonth() + 1) / 3);
            
            // Get the sheet name from the provided URL
            const userSpreadsheet = SpreadsheetApp.openByUrl(sheetUrl); // Open user-provided sheet
            const userSheetName = userSpreadsheet.getName(); // Get the name of the user's sheet

            const emailBody = `
    Hi Team,<br><br>

    Kindly refer to the attached Macroscope FP analysis quarterly report for Q${quarter} ${currentYear}.<br><br>

    Macroscope UI Link: Refer to LookerStudio data studio has security dashboard <a href="${SECURITY_DASHBOARD_URL}">HPS Security Dashboard</a><br>

    Direct Report Link: <a href="${sheetUrl}">${userSheetName} Report</a><br><br>

    Request you to create an action plan accordingly to remediate the vulnerabilities listed by prioritizing critical ones first and acknowledge this mail with further updates.<br><br>

    Just for references, SLA & report data for these vulnerabilities based on the severity is defined as below:<br>

    <div style="max-width: 600px; margin: auto;"> <!-- Limit table width -->
      <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">
        <tr>
          <th style="background-color: lightblue; padding: 8px;">Severity</th>
          <th style="background-color: lightblue; padding: 8px; width: 150px;">Remediation Time</th> <!-- Fixed width for the second column -->
        </tr>
        <tr>
          <td style="border: 1px solid black; padding: 8px;">Critical</td>
          <td style="border: 1px solid black; padding: 8px;">30 days</td>
        </tr>
        <tr>
          <td style="border: 1px solid black; padding: 8px;">High</td>
          <td style="border: 1px solid black; padding: 8px;">60 days</td>
        </tr>
        <tr>
          <td style="border: 1px solid black; padding: 8px;">Medium</td>
          <td style="border: 1px solid black; padding: 8px;">90 days</td>
        </tr>
        <tr>
          <td style="border: 1px solid black; padding: 8px;">Low</td>
          <td style="border: 1px solid black; padding: 8px;">120 days</td>
        </tr>
      </table>
    </div><br><br>

    Do let us know in case of any queries.<br><br>

    Thanks and Regards,<br>
    Security Team
`;



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
        htmlBody: emailDetails.body // Use htmlBody for HTML content
    });

    return true;
}
