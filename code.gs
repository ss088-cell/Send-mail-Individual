// Hardcoded Google Sheet ID for the recipient information (replace with your actual sheet ID)
const RECIPIENT_SHEET_ID = 'your-hardcoded-sheet-id-here';  // Replace with actual sheet ID

// Fixed URL for HPS Security Dashboard
const SECURITY_DASHBOARD_URL = 'https://datastudio.google.com/u/0/reporting/your-dashboard-link-here';

// Function to extract appName and teamName from the Google Sheet name
function extractAppAndTeamName(sheetUrl) {
    try {
        const sheet = SpreadsheetApp.openByUrl(sheetUrl);  // Use openByUrl to handle the full URL
        const sheetName = sheet.getName();

        // Split based on '-' and extract appName and teamName
        const parts = sheetName.split('-');
        if (parts.length >= 6 && parts[0] === "Macroscope Scan") {
            const teamName = parts[1]; // teamName (second part)
            const appName = parts[2]; // appName (third part)
            return { teamName, appName }; // Return both as an object
        }

        return null;
    } catch (error) {
        Logger.log('Error extracting app and team name: ' + error.message);
        return null;
    }
}

// Function to fetch email details based on appName (using hardcoded sheet ID)
function fetchEmailDetails(sheetUrl) {
    const names = extractAppAndTeamName(sheetUrl);
    if (!names) {
        return null;
    }

    const { teamName, appName } = names;

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

                <div style="margin: 0;"> <!-- Remove max-width to stick it to the left -->
                    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: auto; margin: 0;">
                        <tr>
                            <th style="background-color: lightblue; padding: 4px; width: 80px;">Severity</th>
                            <th style="background-color: lightblue; padding: 4px; width: 120px;">Remediation Time</th> <!-- Increased width for header -->
                        </tr>
                        <tr>
                            <td style="border: 1px solid black; padding: 4px;">Critical</td>
                            <td style="border: 1px solid black; padding: 4px;">30 days</td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid black; padding: 4px;">High</td>
                            <td style="border: 1px solid black; padding: 4px;">60 days</td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid black; padding: 4px;">Medium</td>
                            <td style="border: 1px solid black; padding: 4px;">90 days</td>
                        </tr>
                        <tr>
                            <td style="border: 1px solid black; padding: 4px;">Low</td>
                            <td style="border: 1px solid black; padding: 4px;">120 days</td>
                        </tr>
                    </table>
                </div><br><br>

                Do let us know in case of any queries.<br><br>

                Thanks and Regards,<br>
                Security Team
            `;

            // Format subject line using the report name
            const reportDate = new Date();
            const reportDay = reportDate.getDate();
            const reportMonth = reportDate.getMonth() + 1; // Month is 0-indexed
            const reportYear = reportDate.getFullYear();

            const subject = `Macroscope Scan - ${teamName} - ${appName} - ${reportDay} - ${reportMonth} - ${reportYear}`;

            emailDetails = {
                to: data[i][1],
                cc: data[i][2],
                subject: subject,
                body: emailBody
            };

            // Get the folder ID from the recipient data sheet (4th column)
            const folderId = data[i][3]; 
            const fileName = userSheetName; // Name as the original sheet

            // Save the file in the specified folder without converting to PDF
            const file = DriveApp.getFileById(userSpreadsheet.getId()); // Get the original Google Sheet file
            const folder = DriveApp.getFolderById(folderId); // Get the folder where it will be saved
            file.makeCopy(fileName, folder); // Create a copy in the specified folder with the original name

            break;
        }
    }

    return emailDetails;
}

// Function to send the report email from DL
function sendReportEmail(sheetUrl, emailDetails) {
    if (!emailDetails) {
        return false;
    }

    // Change this to your Distribution List email
    const DL_EMAIL = "your-dl-email@example.com"; 

    MailApp.sendEmail({
        to: emailDetails.to,
        cc: emailDetails.cc,
        subject: emailDetails.subject,
        htmlBody: emailDetails.body, // Use htmlBody for HTML content
        replyTo: DL_EMAIL // Ensures replies are sent to the DL
    });

    return true;
}ailDetails.subject,
        htmlBody: emailDetails.body // Use htmlBody for HTML content
    });

    return true;
}
