Public Sub SendEmail(strRecipients As String, strSubject As String, strBody As String, _
    Optional strCC As String, Optional strAttachments As String)

    Dim aOutlook As Object
    Dim aEmail As Object

    Set aOutlook = CreateObject("Outlook.Application")
    Set aEmail = aOutlook.CreateItem(0)

    'Set Recipient
    aEmail.To = strRecipients
    
    'Set CC
    If strCC <> "" Then
        aEmail.CC = strCC
    End If
    
    'Set Subject
    aEmail.Subject = strSubject
    
    'Set Body
    aEmail.body = strBody

    'Set Attachments
    If strAttachments <> "" Then
        aEmail.ATTACHMENTS.Add strAttachments
    End If
    
    'Display for user
    aEmail.display
    
End Sub
