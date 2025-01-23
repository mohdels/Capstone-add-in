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
    Dim ns As Object
    Dim objMail As Object
    Dim categories As Outlook.categories
    Dim firstCategory As Outlook.category
    Dim entryIDs() As String
    Dim i As Integer
    
    ' Split the EntryID collection into individual EntryIDs (in case more than one email comes in at once)
    entryIDs = Split(EntryIDCollection, ",")
    
    ' Get MAPI namespace (used to access Outlook data)
    Set ns = Application.GetNamespace("MAPI")
    
    ' Retrieve the list of categories defined in Outlook
    Set categories = Application.Session.categories

    ' Loop through each EntryID in the collection (using a For loop)
    For i = 0 To UBound(entryIDs)
        ' Get the email item using its EntryID
        Set objMail = ns.GetItemFromID(entryIDs(i))
        
        ' Make sure the item is a MailItem (ignore other types like calendar items)
        If TypeName(objMail) = "MailItem" Then
            ' Check if there are any categories
            If categories.Count > 0 Then
                ' Get the first category from the list
                Set firstCategory = categories.item(1) ' Access the first category
                objMail.categories = firstCategory.Name ' Assign the first category to the email
                objMail.Save ' Save the changes

                ' Optional: Show the category that was assigned (for debugging)
                ' MsgBox "Assigned category: " & firstCategory.Name
            End If
        End If
    Next i
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


