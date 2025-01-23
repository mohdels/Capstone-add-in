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
    Dim objMail As Outlook.MailItem
    Dim entryIDs() As String
    Dim i As Long
    Dim csvFilePath As String
    Dim fso As Object
    Dim csvFile As Object
    Dim csvLine As String

    ' Set the path for the CSV file relative to the user's Documents folder
    csvFilePath = Environ("USERPROFILE") & "\Documents\Emails.csv"

    ' Split the EntryIDCollection into an array
    entryIDs = Split(EntryIDCollection, ",")

    ' Get the Outlook namespace
    Set ns = Application.GetNamespace("MAPI")

    ' Loop through each EntryID in the collection
    For i = LBound(entryIDs) To UBound(entryIDs)
        ' Get the email item using its EntryID
        Set objMail = ns.GetItemFromID(entryIDs(i))

        ' Make sure the item is a MailItem (ignore other types like calendar items)
        If TypeName(objMail) = "MailItem" Then
            ' Open or append to the CSV file
            On Error Resume Next
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FileExists(csvFilePath) Then
                On Error GoTo HandleFileError
                Set csvFile = fso.OpenTextFile(csvFilePath, 8) ' Append mode
            Else
                Set csvFile = fso.CreateTextFile(csvFilePath, True)
                ' Write the CSV header
                csvFile.WriteLine "Date Received,Subject,Sender,Has Attachments,Categories"
            End If

            ' Write the email details to the CSV file
            csvLine = """" & objMail.ReceivedTime & """," & _
                      """" & objMail.Subject & """," & _
                      """" & objMail.SenderName & """," & _
                      """" & IIf(objMail.Attachments.Count > 0, "Yes", "No") & """," & _
                      """" & objMail.categories & """"
            csvFile.WriteLine csvLine
            csvFile.Close
        End If
    Next i

    Exit Sub

HandleFileError:
    MsgBox "Error: Unable to access the CSV file. Please ensure it is not open or locked by another application.", vbExclamation
    Exit Sub
End Sub

Sub GetEmailsWithCategory()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim items As Outlook.items
    Dim filteredItems As Outlook.items
    Dim item As Object
    Dim category As String
    Dim i As Integer

    ' Define the category to search for
    category = "fun" ' Change this to your desired category name

    ' Get the MAPI namespace
    Set ns = Application.GetNamespace("MAPI")

    ' Get the Inbox folder
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Get all items in the Inbox folder
    Set items = inbox.items

    ' Restrict the items collection to only those with the specific category
    Set filteredItems = items.Restrict("[Categories] = '" & category & "'")

    ' Loop through the filtered items and display information
    For i = 1 To filteredItems.Count
        Set item = filteredItems.item(i)
        ' Example: Display the Subject of each email in the filtered items
        MsgBox "Subject: " & item.Subject
    Next i
End Sub





