Public Function ParseWords(ByRef Text As String, ByRef arrKeywords() As Variant) As String

    Dim strParse As String
    Dim arr() As String
    Dim strRegexHTML As String
    
    
    strParse = LCase(Trim(Text))
    strParse = WriteReadText(strParse)
    strParse = RemovePunctuation(strParse)
    strParse = RemoveStopWords(strParse)
    strParse = RemoveNumbers(strParse)
    strParse = Trim(strParse)
    strParse = RemoveByArray(strParse, arrKeywords)
    strParse = SplitAndProcess(strParse)
    strParse = RemoveSpaces(strParse)
    strParse = RemoveDuplicateWords(strParse)
    strParse = RemoveSpaces(strParse)
    ParseWords = strParse

    
End Function

Public Function WriteReadText(Text As String) As String

    Dim fso As Object
    Dim strPath As String
    Dim openApp As New clsXLApp
    Dim objTextStream
    
    Set openApp.myXLApp = Application
    strPath = openApp.DownloadFolder & "parse.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(strPath)
    oFile.WriteLine Text
    oFile.Close
    Set objTextStream = fso.OpenTextFile(strPath, 1)
    WriteReadText = objTextStream.ReadAll
    objTextStream.Close
    
    Set fso = Nothing
    Set oFile = Nothing
    Set openApp = Nothing
    Set objTextStream = Nothing
    
    Kill strPath
    
End Function

Public Function SplitAndProcess(Text As String) As String
    
    On Error Resume Next
    
    Dim strText As String
    Dim arrSplit() As String
    Dim arrClean() As String
    Dim arrFinal() As String
    Dim i As Integer
    Dim j As Integer
    Dim strArrayText As String
    Dim strOutputText As String

    strText = Text
    
    arrSplit = Split(strText, " ", , vbBinaryCompare)
    
    For i = 0 To UBound(arrSplit)
        strArrayText = arrSplit(i)
        strArrayText = Trim(strArrayText)
        arrSplit(i) = strArrayText
    Next i
    
    arrClean = RemoveBlankArrayItems(arrSplit)
    

    arrFinal = GetPOSFromArray(arrClean)
    
    
    For j = 0 To UBound(arrFinal)
        strArrayText = arrFinal(j)
        strArrayText = Trim(strArrayText)
        If j = 0 Then
            strOutputText = strArrayText
        Else
            strOutputText = strOutputText & ", " & strArrayText
        End If
    Next j
    
    SplitAndProcess = strOutputText
    
End Function

Public Function RemoveBlankArrayItems(ByRef ArrayWithBlanks() As String) As String()
    
    On Error Resume Next
    
    Dim base As Long
    base = LBound(ArrayWithBlanks)

    Dim result() As String
    ReDim result(base To UBound(ArrayWithBlanks))

    Dim countOfNonBlanks As Long
    Dim i As Long
    Dim myElement As String

    For i = base To UBound(ArrayWithBlanks)
        myElement = ArrayWithBlanks(i)
        If myElement <> vbNullString Then
            result(base + countOfNonBlanks) = myElement
            countOfNonBlanks = countOfNonBlanks + 1
        End If
    Next i
    If countOfNonBlanks = 0 Then
        ReDim result(base To base)
    Else
        ReDim Preserve result(base To base + countOfNonBlanks - 1)
    End If

    RemoveBlankArrayItems = result

    
End Function

Public Function RemoveSpaces(Text As String) As String

    Dim i As Integer
    Dim strText As String
    Dim strCheck As String
    
     
    strText = Text 'define string
    
    For i = 1 To Len(strText)
        strCheck = Mid(strText, i, 1)
        
        If strCheck = " " Then
            strText = Replace(strText, i, "", 1, , vbBinaryCompare)
        End If
    Next
    
    RemoveSpaces = strText
    
End Function

Public Function RemoveNumbers(Text As String) As String

    Dim i As Integer
    Dim strText As String
    Dim strCurrent As String
    
    strText = Text 'define string
    
    For i = 1 To Len(strText)
        strCurrent = Mid(strText, i, 1)
        
        If IsNumeric(strCurrent) = True Then
            strText = Replace(strText, Mid(strText, i, 1), " ", 1, , vbBinaryCompare)
        End If
    Next
    
    RemoveNumbers = strText
    
End Function



Public Function GetPOSFromArray(ByRef arr() As String) As String()
  On Error Resume Next
  
  Dim objWord As Word.Application
  Dim mySynInfo As Word.SynonymInfo
  Dim varList As Variant
  Dim varPos As Variant
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim iMax As Integer
  Dim strPos As String
  Dim strArray() As String
  Dim strFinal() As String
 
  Set objWord = CreateObject("Word.Application")
  
  iMax = 1
    
    ReDim strArray(UBound(arr), 1)
 
    For j = 0 To UBound(arr)
        Set mySynInfo = SynonymInfo(Word:=arr(j), LanguageID:=wdEnglishUS)
        If mySynInfo.MeaningCount <> 0 Then
            varPos = mySynInfo.PartOfSpeechList
            If i > iMax Then iMax = i
                For i = 1 To UBound(varPos)
                  Select Case varPos(i)
                    Case wdAdjective
                      strPos = "adjective"
                    Case wdNoun
                      strPos = "noun"
                    Case wdAdverb
                      strPos = "adverb"
                    Case wdVerb
                      strPos = "verb"
                    Case wdConjunction
                      strPos = "conjunction"
                    Case wdIdiom
                      strPos = "idiom"
                    Case wdInterjection
                      strPos = "interjection"
                    Case wdPreposition
                      strPos = "preposition"
                    Case wdPronoun
                      strPos = "pronoun"
                     Case Else
                      strPos = "other"
                  End Select
                    strArray(j, 0) = Trim(arr(j))
                    strArray(j, 1) = strPos
                Next i
            Else
                strArray(j, 0) = arr(j)
                strArray(j, 1) = "none"
            End If

    Next j
    
    
    k = 0
    
    
    For i = 0 To UBound(strArray)
        If strArray(i, 1) <> "none" And strArray(i, 1) <> "conjunction" And strArray(i, 1) <> "interjection" And strArray(i, 1) <> "preposition" And strArray(i, 1) <> "pronoun" And strArray(i, 1) <> "adverb" And strArray(i, 0) <> vbNullString Then
            ReDim Preserve strFinal(k)
            strFinal(k) = strArray(i, 0)
            k = k + 1
        End If
    Next i
    
    GetPOSFromArray = strFinal
    objWord.Quit
    Set objWord = Nothing
    
End Function


Public Function RemoveByArray(ByRef Text As String, ByVal arr As Variant) As String

    Dim item As Variant
    Dim strStrip As String
    
    strStrip = Text
    

    For Each item In arr
        strStrip = Replace(strStrip, item, " ", 1, , vbBinaryCompare)
    Next item
 
    RemoveByArray = strStrip

End Function

Public Function RemovePunctuation(ByVal Text As String) As String

    Dim arrPunctuation() As Variant
    Dim strParse As String
    
    arrPunctuation = Array("!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "-", "_", "+", "=", "{", "}", "[", "]", "|", "\", ":", """", ";", "'", "<", ",", ">", ".", "/", "?", "`", "~", "–")
    
    strParse = Text
    
    For Each item In arrPunctuation
        strParse = Replace(strParse, item, " ", 1, , vbBinaryCompare)
    Next item
    
    RemovePunctuation = strParse
    
End Function


Public Function RemoveStopWords(ByVal Text As String) As String
    
    Dim arrStop() As Variant
    Dim strParse As String
    Dim i As Integer
    Dim lrow As Range
    Dim tbl As ListObject
    
    Set tbl = ThisWorkbook.Worksheets("SW").ListObjects("StopWords")
    
    i = 0
    
    For Each lrow In tbl.ListColumns("Word").DataBodyRange.Rows
        ReDim Preserve arrStop(i)
        arrStop(i) = " " & Trim(lrow.Value2) & " "
        i = i + 1
    Next lrow
    
    strParse = LCase(Trim(Text))
    
    strParse = RemoveByArray(strParse, arrStop)
    
    RemoveStopWords = strParse
    
End Function

Public Function RegexReplace(Text_Pattern As String, Text_Input As String, Text_Replace As String) As String
    
    Dim regexObject As RegExp
    Set regexObject = New RegExp
    
    With regexObject
        .pattern = Text_Pattern
    End With

    RegexReplace = regexObject.Replace(Text_Input, Text_Replace)
    
    
End Function


Public Function RemoveDuplicateWords(Text As String) As String

    Dim str
    Dim i As Long
    str = Split(Text, ", ")
    For i = 0 To UBound(str)
        If InStr(1, RemoveDuplicateWords, str(i), 1) = 0 Then RemoveDuplicateWords = Trim(RemoveDuplicateWords) & ", " & str(i)
    Next i
    RemoveDuplicateWords = Trim(RemoveDuplicateWords)
    
End Function



Private Function ReDimPreserve(ByRef avarArrayToPreserve As Variant, ByVal varNewFirstUBound As Variant, _
    ByVal varNewLastUBound As Variant) As Variant
'Workaround for Redim Preserve issue with multidimensional arrays
Dim lngOldFirstUBound As Long
Dim lngOldLastUBound As Long
Dim lngFirst As Long
Dim lngLast As Long
Dim avarPreservedArray As Variant

    ReDimPreserve = False
    'check if its in array first
    If IsArray(avarArrayToPreserve) Then
        'create new array
        ReDim avarPreservedArray(varNewFirstUBound, varNewLastUBound)
        'get old lBound/uBound
        lngOldFirstUBound = UBound(avarArrayToPreserve, 1)
        lngOldLastUBound = UBound(avarArrayToPreserve, 2)
        'loop through first
        For lngFirst = LBound(avarArrayToPreserve, 1) To varNewFirstUBound
            For lngLast = LBound(avarArrayToPreserve, 2) To varNewLastUBound
                'if its in range, then append to new array the same way
                If lngOldFirstUBound >= lngFirst And lngOldLastUBound >= lngLast Then
                    avarPreservedArray(lngFirst, lngLast) = avarArrayToPreserve(lngFirst, lngLast)
                End If
            Next
        Next
        'return the array redimmed
        If IsArray(avarPreservedArray) Then ReDimPreserve = avarPreservedArray
    End If
End Function
