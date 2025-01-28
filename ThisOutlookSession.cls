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
    Dim items As Outlook.items
    Dim filteredItems As Outlook.items
    Dim filteredItem As Object
    Dim csvFilePath As String
    Dim fso As Object
    Dim csvFile As Object
    Dim csvLine As String
    Dim recipient As Outlook.recipient
    Dim recipientEmails As String
    Dim customFolderPath As String
    
    ' Define the path for the custom folder in ProgramData
    customFolderPath = Environ("PROGRAMDATA") & "\MyApp\"
    
    ' Ensure the folder exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(customFolderPath) Then
        fso.CreateFolder customFolderPath
    End If
    
    ' Set the full path for the CSV file
    csvFilePath = customFolderPath & "EmailsFrom2025.csv"

    ' Get the MAPI namespace and Inbox folder
    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Get all items in the Inbox folder
    Set items = inbox.items

    ' Filter emails received from 2025 onwards
    Set filteredItems = items.Restrict("[ReceivedTime] >= '01/01/2025'")

    ' Check if filteredItems contains any emails
    If filteredItems.Count = 0 Then
        MsgBox "No emails found from 2025 onward.", vbExclamation
        Exit Sub
    End If

    ' Open or create the CSV file
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(csvFilePath) Then
        Set csvFile = fso.OpenTextFile(csvFilePath, 8) ' Append mode
    Else
        Set csvFile = fso.CreateTextFile(csvFilePath, True)
        ' Write the CSV header
        csvFile.WriteLine "senderEmail,receiverEmail,date,subject,body,category,assignedTo,hasAttachments"
    End If

    ' Loop through the filtered items and write their details to the CSV
    For Each filteredItem In filteredItems
        If TypeName(filteredItem) = "MailItem" Then
            ' Initialize recipientEmails
            recipientEmails = ""
    
            ' Collect recipient email addresses
            On Error Resume Next
            If Not filteredItem.Recipients Is Nothing Then
                For Each recipient In filteredItem.Recipients
                    If Not recipient Is Nothing Then
                        recipientEmails = recipientEmails & recipient.Address & ";"
                    End If
                Next recipient
            End If
            On Error GoTo 0
    
            ' Remove the trailing semicolon
            If Len(recipientEmails) > 0 Then
                recipientEmails = Left(recipientEmails, Len(recipientEmails) - 1)
            End If
    
            ' Write the email details to the CSV
            csvLine = """" & filteredItem.SenderEmailAddress & """," & _
                      """" & recipientEmails & """," & _
                      """" & filteredItem.ReceivedTime & """," & _
                      """" & filteredItem.Subject & """," & _
                      """" & Replace(filteredItem.Body, vbNewLine, " ") & """," & _
                      """" & filteredItem.categories & """," & _
                      """NotAssigned""" & "," & _
                      """" & IIf(filteredItem.Attachments.Count > 0, "Yes", "No") & """"
            On Error Resume Next
            csvFile.WriteLine csvLine
            On Error GoTo 0
        End If
    Next filteredItem


    ' Close the CSV file
    csvFile.Close

    MsgBox "Emails from 2025 till now have been saved to: " & csvFilePath, vbInformation
End Sub


