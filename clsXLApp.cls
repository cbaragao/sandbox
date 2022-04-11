Option Explicit

Public WithEvents myXLApp As Application


Private Sub myXLApp_WorkbookOpen(ByVal wb As Workbook)
    Dim wks As Worksheet
    Dim wksCount As Integer
    Dim i As Integer
    
    wksCount = wb.Worksheets.Count
    

    For i = 1 To wksCount
        If ActiveWorkbook.Worksheets(i).Name <> "Main" Then
            ActiveWorkbook.Worksheets(i).Visible = xlVeryHidden
        End If
    Next i
    
    With wb
        wb.RefreshAll
    End With
    
    With myXLApp
        .WindowState = xlMaximized
    End With
End Sub

Property Get UserName() As String

    UserName = UCase(Environ("USERNAME"))
    
End Property

Property Get UserComputerName() As String

    ComputerName = UCase(Environ("COMPUTERNAME"))
    
End Property

Property Get UserDomain() As String

    UserDomain = UCase(Environ("USERDOMAIN"))

End Property

Property Get UserProfile() As String

    UserProfile = UCase(Environ("USERPROFILE"))

End Property

Property Get FolderPath() As String

    FolderPath = ThisWorkbook.Path

End Property

Property Get FilePath() As String

    FilePath = ThisWorkbook.FullName
    
End Property

Property Get DownloadFolder() As String

    DownloadFolder = "C:\Users\" & LCase(Environ("USERNAME")) & "\Downloads\"

End Property
