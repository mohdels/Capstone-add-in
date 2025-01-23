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
    Dim objMail As MailItem
    Dim ns As NameSpace
    Dim categoryList As categories
    Dim category As Variant
    Dim responseText As String
    Dim httpRequest As Object
    Dim url As String
    Dim emailData As String
    Dim categories As Variant
    Dim categoryExists As Boolean

    ' Initialize Outlook Namespace
    Set ns = Application.GetNamespace("MAPI")
    
    ' Retrieve the incoming mail item using its Entry ID
    Set objMail = ns.GetItemFromID(EntryIDCollection)
    
    ' Escape special characters and prepare JSON payload
    emailData = CreateJsonString(objMail.Subject, objMail.SenderName, objMail.Recipients(1).Name, objMail.body)
    
    ' Define the URL for the Flask server
    url = "http://localhost:3000/email"
    
    ' Create the HTTP request object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Open a POST request
    httpRequest.Open "POST", url, False
    
    ' Set the Content-Type header
    httpRequest.setRequestHeader "Content-Type", "application/json"
    
    ' Send the request with the email data
    On Error Resume Next
    httpRequest.Send emailData
    If Err.Number <> 0 Then
        MsgBox "HTTP Request Error: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get the response text
    responseText = httpRequest.responseText
    
    ' Parse the JSON response into an array of categories
    categories = ParseJson(responseText)
    
    ' Loop through the categories returned by the backend
    For Each category In categories
        categoryExists = False
        
        ' Check if the category already exists in Outlook
        For Each c In ns.categories
            If c.Name = category Then
                categoryExists = True
                Exit For
            End If
        Next c
        
        ' If the category doesn't exist, create it
        If Not categoryExists Then
            ns.categories.Add category
        End If
        
        ' Apply the category to the email
        If objMail.categories = "" Then
            objMail.categories = category
        Else
            objMail.categories = objMail.categories & "," & category
        End If
    Next category
    
    ' Save changes to the email
    objMail.Save
    
    ' Clean up
    Set httpRequest = Nothing
    Set ns = Nothing
End Sub

' Helper function to escape special characters in JSON
Function EscapeJsonString(value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbTab, "\t")
    value = Replace(value, vbLf, "\n")
    EscapeJsonString = value
End Function

' Function to create a JSON string
Function CreateJsonString(title As String, sender As String, receiver As String, body As String) As String
    Dim json As String
    json = "{""title"":""" & EscapeJsonString(title) & """," & _
           """sender"":""" & EscapeJsonString(sender) & """," & _
           """receiver"":""" & EscapeJsonString(receiver) & """," & _
           """body"":""" & EscapeJsonString(body) & """}"
    CreateJsonString = json
End Function

' Simple function to parse JSON response
Function ParseJson(jsonString As String) As Variant
    Dim result() As String
    Dim i As Integer

    ' Remove brackets from the JSON array
    If Left(jsonString, 1) = "[" And Right(jsonString, 1) = "]" Then
        jsonString = Mid(jsonString, 2, Len(jsonString) - 2)
    End If

    ' Split the string by commas
    result = Split(jsonString, ",")
    
    ' Clean up quotes from the categories
    For i = 0 To UBound(result)
        result(i) = Trim$(Replace$(result(i), """", ""))
    Next i
    
    ParseJson = result
End Function


