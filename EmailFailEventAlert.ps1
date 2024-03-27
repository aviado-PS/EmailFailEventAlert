<#
.SYNOPSIS
This script monitors and alerts administrators about 'FAIL' events detected in the Exchange Server message tracking logs.

.DESCRIPTION
"EmailFailEventAlert.ps1" is designed to proactively monitor the Exchange Server environment for email delivery issues, such as failures due to spam filters, malware detection, or other transport rule-related causes. When a 'FAIL' event is detected within the last 2 hours, the script sends an alert via email with details about the failure. This assists administrators in quickly addressing and resolving email delivery issues.

.AUTHOR
Aviad Ofek

.EXAMPLE
.\EmailFailEventAlert.ps1

This command runs the script to check for 'FAIL' events in the last 2 hours and sends an email alert if any are found.

.NOTES
Requires Exchange Management Shell. Ensure you have the necessary permissions to execute message tracking log searches and send emails through your Exchange Server.

#>

# Define time range
$startDate = (Get-Date).AddHours(-2)  # Checking for the last 2 hours
$endDate = Get-Date

# Sample data for demonstration
$failEvents = @(
    [PSCustomObject]@{
        Timestamp = Get-Date
        Source = "SMTP"
        Sender = "sender@example.com"
        Recipients = "recipient1@example.com; recipient2@example.com"
        MessageSubject = "Sample Subject"
        EventId = "FAIL"
    },
    [PSCustomObject]@{
        Timestamp = (Get-Date).AddHours(-1)
        Source = "SMTP"
        Sender = "another_sender@example.com"
        Recipients = "recipient3@example.com"
        MessageSubject = "Another Sample Subject"
        EventId = "FAIL"
    }
)

# Check if any 'FAIL' events are found and send an alert
if ($failEvents) {
    # Convert the events to HTML format with table styling
    $htmlBody = "<html><body>"
    $htmlBody += "<style>table {border-collapse: collapse; width: 100%;} th, td {border: 1px solid black; padding: 8px; text-align: left;} th {background-color: #4CAF50; color: white;}</style>"
    $htmlBody += $failEvents | ConvertTo-Html -Property Timestamp, Source, Sender, Recipients, MessageSubject, EventId -Title "FAIL Event Details" -PreContent "<h2>Detected FAIL Events</h2>" | Out-String
    $htmlBody += "</body></html>"

    # Sending an email with the HTML body
    Send-MailMessage -From alert@example.com -To admin@example.com -Cc anotheradmin@example.com -Subject "FAIL Event Detected" -Body $htmlBody -BodyAsHtml -SmtpServer "mail.example.com"
}
