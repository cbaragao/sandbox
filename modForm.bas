Option Explicit


Public Sub DeselectListBoxItem(ByRef lb As Control, ByVal Column As Integer, ByVal Text As String)

    Dim i As Integer
    Dim intCounter As Integer
    
    intCounter = 0
    
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) = True Then
            If lb.List(i, Column) = Text Then
                lb.Selected(i) = False
            End If
        End If
    Next i

End Sub

Public Sub SelectListBoxItem(ByRef lb As Control, ByVal Column As Integer, ByVal Text As String)

    Dim i As Integer
    Dim intCounter As Integer
    
    intCounter = 0
    
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) = False Then
            If lb.List(i, Column) = Text Then
                lb.Selected(i) = True
            End If
        End If
    Next i

End Sub

Public Function SumListboxText(ByRef lb As Control, ByVal Column As Integer, ByVal Text As String) As Integer

    Dim i As Integer
    Dim intCounter As Integer
    
    intCounter = 0
    
    For i = 0 To lb.ListCount - 1
        If lb.Selected(i) = True Then
            If lb.List(i, Column) = Text Then
                intCounter = intCounter + 1
            End If
        End If
    Next i
    
    SumListboxText = intCounter
    
End Function


Public Sub SelectLBRows(ByRef lb As Control, ByVal Column As Integer)
    
    Dim i As Integer
    
    For i = 0 To lb.ListCount - 1
        If Trim(lb.List(i, Column)) <> "" Then
            lb.Selected(i) = True
        End If
    Next i

End Sub

Public Sub DeleteEmptyRows(ByRef lb As Control, ByVal Column As Integer)
    
    Dim i As Integer
    
    For i = 0 To lb.ListCount - 1
        If Trim(lb.List(i, Column)) = "" Then
            lb.RemoveItem (i)
        End If
    Next i
    
End Sub

Public Sub CopyTextToClipboard(ByVal Text As String)
'PURPOSE: Copy a given text to the clipboard (using DataObject)
'SOURCE: www.TheSpreadsheetGuru.com
'NOTES: Must enable Forms Library: Checkmark Tools > References > Microsoft Forms 2.0 Object Library
    Dim txt As String
    Dim obj As New DataObject
    txt = Text

If Text = "" Then Exit Sub

'Make object's text equal above string variable
    obj.SetText txt

'Place DataObject's text into the Clipboard
    obj.PutInClipboard


End Sub

Public Sub ListBoxSortDesc(lb As MSForms.ListBox)

'Create variables
Dim i, j, k, l, m As Long
Dim temp() As Variant

ReDim temp(0, 4)

'Use Bubble sort method
    With lb
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                k = CInt(.List(i, 4))
                l = CInt(.List(i + 1, 4))
                If k < l Then
                
                    For m = 0 To 4
                        temp(0, m) = .List(i, m)
                    Next m
                    
                    For m = 0 To 4
                        .List(i, m) = .List(i + 1, m)
                    Next m
                    
                    For m = 0 To 4
                       .List(i + 1, m) = temp(0, m)
                    Next m
                    
                End If
            Next i
        Next j
    End With

End Sub

Public Sub FilterListBox(ByRef lb As Control, ByVal FilterText As String, ByVal KeyColumn As Integer, ByVal TotalColumns As Integer)
    'To avoid any screen update until the process is finished
    Application.ScreenUpdating = False
    'This method must make sure to turn this property back to True before exiting by
    '  always going through the exit_sub label

    On Error GoTo err_sub

    'This will be the string to filter by
    Dim filterSt As String: filterSt = FilterText

    'This is the number of the column to filter by
    Dim filterCol As Integer
    
    filterCol = KeyColumn 'This number can be changed as needed

    'This is the sheet to load the listbox from
    Dim dataSh As Worksheet: Set dataSh = ThisWorkbook.Worksheets("Sheet1") 'The sheet name can be changed as needed

    'This is the number of columns that will be loaded from the sheet (starting with column A)
    Dim colCount As Integer
    
    colCount = TotalColumns 'This constant allows you to easily include more/less columns in future

    'Determining how far down the sheet we must go
    Dim usedRng As Range: Set usedRng = dataSh.UsedRange
    Dim lastRow As Long: lastRow = usedRng.Row - 1 + usedRng.Rows.Count

    Dim c As Long

    'Getting the total width of all the columns on the sheet
    Dim colsTotWidth As Double: colsTotWidth = 0
    For c = 1 To colCount
        colsTotWidth = colsTotWidth + dataSh.Columns(c).ColumnWidth
    Next

    'Determining the desired total width for all the columns in the listbox
    Dim widthToUse As Double
    'Not sure why, but subtracting 4 ensured that the horizontal scrollbar would not appear
    widthToUse = lb.Width - 4
    If widthToUse < 0 Then widthToUse = 0

    'Reset the listbox
    lb.Clear
    lb.ColumnCount = colCount
'    lb.ColumnWidths = colWidthSt
    lb.ColumnHeads = False

    'Reading the entire data sheet into memory
    Dim dataArr As Variant: dataArr = dataSh.UsedRange
    If Not IsArray(dataArr) Then dataArr = dataSh.Range("A1:A2")

    'If filterCol is beyond the last column in the data sheet, leave the list blank and simply exit
    If filterCol > UBound(dataArr, 2) Then GoTo exit_sub 'Do not use Exit Sub here, since we must turn ScreenUpdating back on

    'This array will store the rows that meet the filter condition
    'NB: This array will store the data in transposed form (rows and columns inverted) so that it can be easily
    '    resized later using ReDim Preserve, which only allows you to resize the last dimension
    ReDim filteredArr(1 To colCount, 1 To UBound(dataArr, 1)) 'Make room for the maximum possible size
    Dim filteredCount As Long: filteredCount = 0

    'Copy the matching rows from [dataArr] to [filteredArr]
    'IMPORTANT ASSUMPTION: The first row on the sheet is a header row
    Dim r As Long
    For r = 2 To lastRow
        'The first row will always be added to give the listbox a header
        If r > 1 And InStr(1, dataArr(r, filterCol) & "", filterSt, vbBinaryCompare) = 0 Then
            GoTo continue_for_r
        End If

        'NB: The Like operator is not used above in case [filterSt] has wildcard characters in it
        '    Also, the filtering above is case-insensitive
        '    (if needed, it can be changed to case-sensitive by changing the last parameter to vbBinaryCompare)

        filteredCount = filteredCount + 1
        For c = 1 To colCount
            'Inverting rows and columns in [filteredArr] in preparation for the later ReDim Preserve
            filteredArr(c, filteredCount) = dataArr(r, c)
        Next
        
    

continue_for_r:
    Next

    'Copy [filteredArr] to the listbox, removing the excess rows first
    If filteredCount > 0 Then
        ReDim Preserve filteredArr(1 To colCount, 1 To filteredCount)
        lb.Column = filteredArr
        'Used .Column instead of .List above, as per advice at
        '  https://stackoverflow.com/questions/54204164/listbox-error-could-not-set-the-list-property-invalid-property-value/54206396#54206396
    End If

exit_sub:
    Application.ScreenUpdating = True
    Exit Sub

err_sub:
    MsgBox "Error " & Err.Number & vbCrLf & vbCrLf & Err.Description
    Resume exit_sub 'To make sure that screen updating is turned back on
End Sub
