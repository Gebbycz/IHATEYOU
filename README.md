What it does:

Destroys user experience:
Swaps left and right mouse buttons.
Changes Windows colors to an ugly black/green theme.
Disables Task Manager, Control Panel, Shutdown, Logoff, CMD, and PowerShell via registry changes.

Harvests email addresses:
Scans Outlook data folder + many common user folders (Desktop, Documents, Downloads, AppData, Temp, etc.).
Looks for emails in a very wide range of file types, including:
.txt, .eml, .msg, .html, .htm, .xml, .json, .csv, .log, .ini, .cfg, .dat, .rtf, .md, .sql, .php, .js, .css, .asp, .aspx, .vcf, .conf, .config, .properties, .yaml, .yml, .bak, .old, and more.
Extracts all email addresses using regex from matching files.

Spreads itself:
Creates Outlook emails to every harvested address (including the victim).
Subject: "Email from: [address]"
Body: Insulting message ("I hate you. Don't talk to me ever again...") + system info.
Attaches the script itself (self-propagation).
Sends the emails.

Logging:
Writes a detailed report to C:\temp\email_scan_results.txt.
Shows a "Complete" message box at the end.
