VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim mail As Outlook.MailItem
    Dim newMail As Object
    Dim csvFilePath As String
    Dim txtFilePath As String
    Dim fso As Object
    Dim csvFile As Object
    Dim txtFile As Object
    Dim csvLine As String
    Dim recipient As Outlook.recipient
    Dim recipientEmails As String
    Dim customFolderPath As String
    Dim senderEmail As String

    ' Define the path for the custom folder in ProgramData
    customFolderPath = Environ("PROGRAMDATA") & "\MyApp\"

    ' Ensure the folder exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(customFolderPath) Then
        fso.CreateFolder customFolderPath
    End If

    ' Set the full paths for the CSV and TXT files
    csvFilePath = customFolderPath & "EmailsLog.csv"
    txtFilePath = customFolderPath & "latest_email.txt"

    ' Get the MAPI namespace
    Set ns = Application.GetNamespace("MAPI")

    ' Get the new email using the EntryID
    Set newMail = ns.GetItemFromID(EntryIDCollection)

    ' Ensure it is a MailItem
    If TypeName(newMail) <> "MailItem" Then Exit Sub
    Set mail = newMail

    ' ? Get the real SMTP email for the sender using PropertyAccessor
    senderEmail = GetSMTPFromProperty(mail)

    ' Initialize recipientEmails
    recipientEmails = ""

    ' Collect recipient SMTP email addresses
    On Error Resume Next
    If Not mail.Recipients Is Nothing Then
        For Each recipient In mail.Recipients
            If Not recipient Is Nothing Then
                recipientEmails = recipientEmails & GetSMTPAddress(recipient) & ";"
            End If
        Next recipient
    End If
    On Error GoTo 0

    ' Remove the trailing semicolon
    If Len(recipientEmails) > 0 Then
        recipientEmails = Left(recipientEmails, Len(recipientEmails) - 1)
    End If

    ' Open or create the CSV file
    If fso.FileExists(csvFilePath) Then
        Set csvFile = fso.OpenTextFile(csvFilePath, 8) ' Append mode
    Else
        Set csvFile = fso.CreateTextFile(csvFilePath, True)
        ' Write the CSV header
        csvFile.WriteLine "senderEmail,receiverEmail,date,subject,body,category,assignedTo,hasAttachments"
    End If

    ' Write the email details to the CSV
    csvLine = """" & senderEmail & """," & _
              """" & recipientEmails & """," & _
              """" & mail.ReceivedTime & """," & _
              """" & mail.Subject & """," & _
              """" & Replace(mail.Body, vbNewLine, " ") & """," & _
              """" & mail.categories & """," & _
              """NotAssigned""" & "," & _
              """" & IIf(mail.Attachments.Count > 0, "Yes", "No") & """"
    On Error Resume Next
    csvFile.WriteLine csvLine
    csvFile.Close
    On Error GoTo 0

    ' Update the latest email date/time in the TXT file
    Set txtFile = fso.CreateTextFile(txtFilePath, True)
    txtFile.WriteLine "Latest email received on: " & mail.ReceivedTime
    txtFile.Close

End Sub

' ? Function to retrieve the correct SMTP sender email using PropertyAccessor
Function GetSMTPFromProperty(mail As Outlook.MailItem) As String
    Dim prop As String
    Dim sender As Outlook.AddressEntry
    Dim pa As Outlook.PropertyAccessor
    
    On Error Resume Next
    Set sender = mail.sender
    If sender Is Nothing Then Exit Function

    ' Use PropertyAccessor to get SMTP address
    Set pa = sender.PropertyAccessor
    prop = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")

    ' If the property exists, return it; otherwise, return the default address
    If prop <> "" Then
        GetSMTPFromProperty = prop
    Else
        GetSMTPFromProperty = mail.SenderEmailAddress
    End If
    On Error GoTo 0
End Function

' ? Function to get the SMTP address for recipients
Function GetSMTPAddress(obj As Object) As String
    Dim olExchUser As Outlook.ExchangeUser
    Dim olExchAddrEntry As Outlook.AddressEntry

    On Error Resume Next
    Set olExchAddrEntry = obj.AddressEntry

    If Not olExchAddrEntry Is Nothing Then
        If olExchAddrEntry.AddressEntryUserType = olExchangeUserAddressEntry Or _
           olExchAddrEntry.AddressEntryUserType = olExchangeRemoteUserAddressEntry Then
            Set olExchUser = olExchAddrEntry.GetExchangeUser
            If Not olExchUser Is Nothing Then
                GetSMTPAddress = olExchUser.PrimarySmtpAddress
                Exit Function
            End If
        End If
    End If
    
    ' Default to normal address if SMTP conversion fails
    GetSMTPAddress = obj.Address
End Function

Public Sub RetrievePastEmailsOnStartup()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim items As Outlook.items
    Dim mail As Outlook.MailItem
    Dim csvFilePath As String
    Dim txtFilePath As String
    Dim fso As Object
    Dim csvFile As Object
    Dim txtFile As Object
    Dim csvLine As String
    Dim recipient As Outlook.recipient
    Dim recipientEmails As String
    Dim customFolderPath As String
    Dim latestTimestamp As Date
    Dim mostRecentEmailDate As Date
    Dim item As Object
    
    ' Define the path for the custom folder in ProgramData
    customFolderPath = Environ("PROGRAMDATA") & "\MyApp\"

    ' Ensure the folder exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(customFolderPath) Then
        fso.CreateFolder customFolderPath
    End If

    ' Set the full paths for the CSV and TXT files
    csvFilePath = customFolderPath & "EmailsLog.csv"
    txtFilePath = customFolderPath & "latest_email.txt"

    ' ? Read the last recorded timestamp from the TXT file
    If fso.FileExists(txtFilePath) Then
        Set txtFile = fso.OpenTextFile(txtFilePath, 1) ' Read mode
        latestTimestamp = CDate(Replace(txtFile.ReadLine, "Latest email received on: ", ""))
        txtFile.Close
    Else
        ' If the file does not exist, assume a very old date
        latestTimestamp = #1/1/2000# ' Default fallback
    End If

    ' Get the MAPI namespace and Inbox folder
    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Get all emails received after the last timestamp
    Set items = inbox.items
    items.Sort "[ReceivedTime]", True ' Sort newest to oldest
    
    ' ? Loop through emails and log only those received after latestTimestamp
    mostRecentEmailDate = latestTimestamp
    For Each item In items
        If TypeName(item) = "MailItem" Then
            Set mail = item
            
            ' Skip emails before the last recorded timestamp
            If mail.ReceivedTime <= latestTimestamp Then Exit For
            
            ' Track the most recent email's date
            If mail.ReceivedTime > mostRecentEmailDate Then
                mostRecentEmailDate = mail.ReceivedTime
            End If
            
            ' Initialize recipientEmails
            recipientEmails = ""

            ' Collect recipient SMTP email addresses
            On Error Resume Next
            If Not mail.Recipients Is Nothing Then
                For Each recipient In mail.Recipients
                    If Not recipient Is Nothing Then
                        recipientEmails = recipientEmails & GetSMTPAddress(recipient) & ";"
                    End If
                Next recipient
            End If
            On Error GoTo 0

            ' Remove the trailing semicolon
            If Len(recipientEmails) > 0 Then
                recipientEmails = Left(recipientEmails, Len(recipientEmails) - 1)
            End If

            ' Open or create the CSV file
            If fso.FileExists(csvFilePath) Then
                Set csvFile = fso.OpenTextFile(csvFilePath, 8) ' Append mode
            Else
                Set csvFile = fso.CreateTextFile(csvFilePath, True)
                ' Write the CSV header
                csvFile.WriteLine "senderEmail,receiverEmail,date,subject,body,category,assignedTo,hasAttachments"
            End If

            ' ? Get the real SMTP sender email
            senderEmail = GetSMTPFromProperty(mail)

            ' Write the email details to the CSV
            csvLine = """" & senderEmail & """," & _
                      """" & recipientEmails & """," & _
                      """" & mail.ReceivedTime & """," & _
                      """" & mail.Subject & """," & _
                      """" & Replace(mail.Body, vbNewLine, " ") & """," & _
                      """" & mail.categories & """," & _
                      """NotAssigned""" & "," & _
                      """" & IIf(mail.Attachments.Count > 0, "Yes", "No") & """"
            On Error Resume Next
            csvFile.WriteLine csvLine
            csvFile.Close
            On Error GoTo 0
        End If
    Next item

    ' ? Update the latest email date/time in the TXT file
    If mostRecentEmailDate > latestTimestamp Then
        Set txtFile = fso.CreateTextFile(txtFilePath, True)
        txtFile.WriteLine "Latest email received on: " & mostRecentEmailDate
        txtFile.Close
    End If

    MsgBox "Emails retrieved from " & latestTimestamp & " onward have been logged.", vbInformation
End Sub

Private Sub Application_Startup()
    RetrievePastEmailsOnStartup
End Sub

