Option Explicit

Public Sub AddItemToSharePointList(ByVal ListName As String, ByVal SharepointUrl As String, ByVal strBatch As String, _
    Optional strSuccessMessage As String = "List update successfull!", Optional ByVal blnSurpressMessages As Boolean, Optional blnNotDrcExclusionReasons As Boolean = True)
'Adds item to specified SharePoint list
Dim objXMLHTTP As MSXML2.XMLHTTP60
Dim strSoapBody As String
Dim strParseXML As String
On Error GoTo errhandler

    Set objXMLHTTP = New MSXML2.XMLHTTP60
    
    'Creates soap body string based on the xml provided to the procedure
    strSoapBody = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope " & _
        "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
        "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
        "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body><UpdateListItems xmlns=""http://schemas.microsoft.com/sharepoint/soap/""><listName>" & ListName _
        & "</listName><updates>" & strBatch & "</updates></UpdateListItems></soap:Body></soap:Envelope>"
    
    'Contructs http post request and sends it with the soap body
    With objXMLHTTP
        .Open "POST", SharepointUrl + "_vti_bin/Lists.asmx", False ', "USERNAME", "PASSWORD"
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"
        .Send strSoapBody
    End With

    'Checks if a update is successful.
    If objXMLHTTP.Status = 200 Then
        strParseXML = objXMLHTTP.responseText
        strParseXML = Mid(strParseXML, InStr(strParseXML, "<ErrorCode>") + 11, 10)
    Else
        MsgBox objXMLHTTP.responseText
    End If
  
errhandler:
    'Error handling and routine termination
    If blnSurpressMessages = False Then
        Select Case Err.Number
        
        Case Is = 0: 'No error - do nothing
            If strParseXML = "0x00000000" Then
                MsgBox strSuccessMessage, vbOKOnly + vbExclamation, "Success"
            Else
                MsgBox objXMLHTTP.responseText
                MsgBox (Err.Number & " error occurred " & Err.Description)
            End If
        Case Else: 'Unanticipated errors
            MsgBox (Err.Number & " error occurred " & Err.Description)
        End Select
    End If

    Set objXMLHTTP = Nothing
End Sub

Public Function GetSharePointListData(ByVal strListName As String, ByVal strSharepointUrl As String) As String
    Dim objXMLHTTP As XMLHTTP60
    Dim strListNameOrGuid As String
    Dim strSoapBody As String

On Error GoTo errhandler
    Set objXMLHTTP = New XMLHTTP60

    strSoapBody = "<?xml version='1.0' encoding='utf-8'?><soap:Envelope " _
        + "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' " _
        + "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'><soap:Body><GetListItems xmlns=" _
        + "'http://schemas.microsoft.com/sharepoint/soap/'><listName>" & strListName _
        & "</listName><rowLimit>600</rowLimit></GetListItems></soap:Body></soap:Envelope>"
    
    With objXMLHTTP
        .Open "POST", strSharepointUrl + "_vti_bin/Lists.asmx", False
        .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
        .setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
        .Send strSoapBody

        If .Status = 200 Then
            GetSharePointListData = .responseXML.XML
        Else
            MsgBox .Status & .statusText
             Debug.Print "SOAP Response not OK"
        End If
    End With
  
errhandler:
    'Error handling and routine termination
    Select Case Err.Number
        Case Is <> 0:
            If MsgBox("An occured while attempting to retrieve the data. If you would like to make another attempt at " _
               + "retrieving the data click 'Yes'. Please ensure you are connected to the VA network prior to clicking " _
               + "'Yes'. Otherwise, click 'No' to manually enter the data on the next screen.", _
               vbCritical + vbYesNo) = vbYes Then
                Err.Clear
                Call GetSharePointListData(strListName, strSharepointUrl)
            Else
                GetSharePointListData = vbNullString
            End If
    End Select
    Set objXMLHTTP = Nothing
End Function

Public Function CreateXMLString(ByVal UpdateMethod As String, ByRef arrUpdate() As Variant) As String
    
    Dim i As Integer
    Dim strXML As String
    
    
    strXML = "<Batch OnError='Continue'><Method ID='1' Cmd='" & UpdateMethod & "'>"
    
    For i = LBound(arrUpdate) To UBound(arrUpdate)
    
        strXML = strXML & "<Field Name='" & arrUpdate(i, 0) & "'>" & arrUpdate(i, 1) & "</Field>"
    
    Next i
    
        strXML = strXML & "</Method></Batch>"
    
    CreateXMLString = strXML
End Function

'Public Sub ParseRoAddressXml(ByVal strXML As String)
'Dim xDoc As DOMDocument60
'Dim xNodes As IXMLDOMNodeList
'Dim xNodeElement As IXMLDOMElement
'Dim i As Long
'    'Validates that an xml string has been provided to the procedure
'    If strXML <> vbNullString Then
'        Set xDoc = New DOMDocument60
'        xDoc.LoadXML (strXML)
'
'        Call xDoc.SetProperty("SelectionNamespaces", "xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/' xmlns:sp='http://schemas.microsoft.com/sharepoint/soap/' " _
'            + "xmlns:z = '#RowsetSchema' xmlns:rs = 'urn:schemas-microsoft-com:rowset'")
'
'        Set xNodes = xDoc.SelectNodes("//soap:Envelope/soap:Body/sp:GetListItemsResponse/sp:GetListItemsResult/sp:listitems/rs:data/z:row")
'
'        'Resizes array to match list length
'        ReDim varRoAddressData(0 To (xNodes.length - 1), 0 To 7)
'
'        'Populates list into array
'        For Each xNodeElement In xNodes
'            With xNodeElement
'                'Agency Name
'                varRoAddressData(i, 0) = Trim(.getAttribute("ows_LinkTitle"))
'
'                'Address line 1
'                If .getAttribute("ows_ADDRESS1") <> "Null" Then
'                    varRoAddressData(i, 1) = Trim(.getAttribute("ows_ADDRESS1"))
'                End If
'
'                'Address line 2
'                If .getAttribute("ows_ADDRESS2") <> "Null" Then
'                    varRoAddressData(i, 2) = Trim(.getAttribute("ows_ADDRESS2"))
'                End If
'
'                'City
'                If .getAttribute("ows_CITY") <> "Null" Then
'                    varRoAddressData(i, 3) = Trim(.getAttribute("ows_CITY"))
'                End If
'
'                'State
'                If .getAttribute("ows_STATE") <> "Null" Then
'                    varRoAddressData(i, 4) = Trim(.getAttribute("ows_STATE"))
'                End If
'
'                'Zip code
'                If .getAttribute("ows_ZIP") <> "Null" Then
'                    varRoAddressData(i, 5) = Trim(.getAttribute("ows_ZIP"))
'                End If
'            End With
'            i = i + 1
'        Next
'    End If
'End Sub
