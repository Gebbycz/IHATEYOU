Script starts by checking if IHATEYOU.zip exists in the same folder.
Creates Windows Script Host, FileSystem, Outlook, and Regex objects.
Determines its own location on disk.
Sets the email body to:
"I hate you."
"Dont talk to me ever again."
"Pass: 1234" 'Which is the password for the zip
Creates C:\temp if it doesn't already exist.
Creates a log file named email_scan_results.txt.
Attempts to connect to Microsoft Outlook.
Enumerates Outlook folders (Inbox, Sent, Contacts, etc.).
Logs Outlook folder information.
Extracts email addresses from Outlook where possible.
If Outlook isn't available or has few addresses, scans the file system.
Recursively searches these folders:
Documents
Desktop
Downloads
AppData
Temp
Opens supported text-based files.
Searches these file types:
.txt
.log
.csv
.html
.htm
.xml
.eml
.msg
.vcf
.ini
.cfg
.conf
Uses a regular expression to find email addresses.
Removes duplicate addresses.
Creates a new Outlook email for each address found.
Sets the subject to Email from: <recipient>.
Inserts the "I hate you" message into the email body.
Attaches IHATEYOU.zip if it exists.
Otherwise attaches the .vbs script itself.
Sends the email automatically.
Waits about 1.5 seconds before the next email.
Records each successful or failed send in the log.
Saves the final report to C:\temp\email_scan_results.txt.
Displays a popup showing how many emails were processed.
