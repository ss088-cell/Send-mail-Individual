// Hardcoded Google Sheet ID for the recipient information (replace with your actual sheet ID)
const RECIPIENT_SHEET_ID = 'your-hardcoded-sheet-id-here';  // Replace with actual sheet ID

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

// Function to fetch email details with hardcoded email body
function fetchEmailDetails(sheetUrl) {
  const quarter = "Q3"; // Modify based on the current quarter logic
  const year = new Date().getFullYear();

  // Hardcoded email body
  const emailBody = `
Hi Team,

Kindly refer to the attached Macroscope FP analysis quarterly report for ${quarter} ${year}.

Macroscope UI Link: Refer to lookerstudio data studio has security dashboard HPS Security Dashboard (here to HPS Security Dashboard has a link to the dashboard).

Direct Report Link: Name of Google Sheet (here to Name of Google Sheet has a link to the report).

Request you to create an action plan accordingly to remediate the vulnerabilities listed by prioritising critical ones first and acknowledge this mail with further updates.

Just for reference, SLA & report data for these vulnerabilities based on the severity is defined as below:

<table style="border-collapse: collapse; width: 100%;">
  <tr>
    <th style="border: 1px solid black; background-color: lightblue; padding: 8px;">Severity</th>
    <th style="border: 1px solid black; background-color: lightblue; padding: 8px;">Remediation Time</th>
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

Do let us know in case of any queries.

Thanks and Regards,
Security Team
`;

  return {
    to: "recipient@example.com", // Placeholder for email address
    cc: "cc@example.com",         // Placeholder for CC address
    subject: `Macroscope FP Analysis Report for ${quarter} ${year}`,
    body: emailBody
  };
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
