Option Explicit

Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function OCR_PDF(FilePath As String) As String
    Dim objApp As Object
    Dim objPDDoc As Object
    Dim objjso As Object
    Dim wordsCount As Long
    Dim page As Long
    Dim i As Long
    Dim strData As String
    
    
    Set objApp = CreateObject("AcroExch.App")
    Set objPDDoc = CreateObject("AcroExch.PDDoc")
    'AD.1 open file, if =false file is damage
    If objPDDoc.Open(FilePath) Then
        Set objjso = objPDDoc.GetJSObject
        For page = 0 To objPDDoc.GetNumPages - 1
            wordsCount = objjso.GetPageNumWords(page)
            For i = 0 To wordsCount
                'AD.2 Set text to variable strData
                strData = strData & " " & objjso.getPageNthWord(page, i)
            Next i
        Next
        OCR_PDF = strData
    Else
        MsgBox "error!"
    End If
End Function


Public Function OpenFile() As String

    OpenFile = Application.GetOpenFilename(FileFilter:="PDF files (*.pdf*), *.pdf*", Title:="Choose a PDF file to open", MultiSelect:=False)

End Function
