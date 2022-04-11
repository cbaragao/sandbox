Private Const mdatecWWIIStart As Date = "12/07/1941"        'WWII Start Date
Private Const mdatecWWIIEnd As Date = "12/31/1946"          'WWII End Date
Private Const mdatecKorStart As Date = "06/27/1950"         'Korea Start Date
Private Const mdatecKorEnd As Date = "01/31/1955"           'Korea End Date
Private Const mdatecRVNStart As Date = "02/28/1961"         'Vietnam Start Date
Private Const mdatecRVNEnd As Date = "05/07/1975"           'Vietnam End Date
Private Const mdatecGWOTStart As Date = "08/02/1990"        'GWOT Start Date
Private Const mdatecArmyNPRCS As Date = "10/16/1992"        'NPRC Army STR Cutoff
Private Const mdatecNavyNPRCS As Date = "01/31/1994"        'NPRC Navy STR Cutoff
Private Const mdatecUSMCNPRCS As Date = "05/01/1994"        'NPRC Marine STR Cutoff
Private Const mdatecAirForceNPRCS As Date = "05/01/1994"    'NPRC Air Force Reg STR Cutoff
Private Const mdatecAirForceNPRCSG As Date = "06/01/1994"   'NPRC Air Force Guard STR Cutoff
Private Const mdatecCoastNPRCS As Date = "05/01/1998"       'NPRC Coast Guard STR Cutoff
Private Const mdatecArmyNPRCP As Date = "10/01/1994"        'NPRC Army OMPF Cutoff
Private Const mdatecNavyNPRCP As Date = "01/01/1995"        'NPRC Navy OMPF Cutoff
Private Const mdatecAirForceNPRCP As Date = "10/01/2004"    'NPRC Air Force OMPF Cutoff
Private Const mdatecUSMCNPRCP As Date = "01/01/1999"        'NPRC USMC OMPF Cutoff
Private Const mdatecHAIMS As Date = "01/01/2014"            'HAIMS Date
Private Const mdatecHAIMSCG As Date = "09/01/2014"          'HAIMS USCG Date
Private Const mdatecArmyFireSt As Date = "11/01/1912"       'Army Fire-Related Start Date
Private Const mdatecArmyFireEnd As Date = "01/01/1960"      'Army Fire-Related End Date
Private Const mdatecAirForceFireSt As Date = "09/25/1947"   'Air Force Fire-Related Start Date
Private Const mdatecAirForceFireEnd As Date = "01/01/1964"  'Air Force Fire-Related End Date
Private Const mdatecArmyDPRIS As Date = "10/01/1994"        'Army DPRIS
Private Const mdatecNavyDPRIS As Date = "01/01/1995"        'Navy DPRIS
Private Const mdatecAirForceDPRIS As Date = "10/01/2004"    'AF DPRIS
Private Const mdatecMarinesDPRIS As Date = "01/01/1999"     'USMC DPRIS


'***************************************************************
'Controls/Form Info Passed as Parameters
'***************************************************************

Public Function CalcWartime(lb As MSForms.ListBox) As Integer
    
    Dim dateStart As Date
    Dim dateEnd As Date
    Dim intWar As Integer
    Dim i As Integer
    
    intWar = 0

    For i = 0 To lb.ListCount - 1
    dateStart = CDate(lb.List(i, 1))
    dateEnd = CDate(lb.List(i, 2))
        
        For j = dateStart To dateEnd
            If j >= mdatecWWIIStart And j <= mdatecWWIIEnd Then
                intWar = intWar + 1
            End If
            
            If j >= mdatecKorStart And j <= mdatecKorEnd Then
                intWar = intWar + 1
            End If
            
            If j >= mdatecRVNStart And j <= mdatecRVNEnd Then
                intWar = intWar + 1
            End If
            
            If j >= mdatecGWOTStart Then
                intWar = intWar + 1
            End If
        Next
    Next

    CalcWartime = intWar

End Function

Public Function CalcWWII(lb As MSForms.ListBox) As Boolean
        Dim i As Integer
        Dim j As Date
        Dim dateStart As Date
        Dim dateEnd As Date
        
        For i = 0 To lb.ListCount - 1
            For j = dateStart To dateEnd
                If j >= mdatecWWIIStart And j <= mdatecWWIIEnd Then
                    CalcWWII = True
                End If
            Next
        Next
End Function

Public Function CalcKorea(lb As MSForms.ListBox) As Boolean
        Dim i As Integer
        Dim j As Date
        Dim dateStart As Date
        Dim dateEnd As Date
        
        For i = 0 To lb.ListCount - 1
            dateStart = lb.List(i, 1)
            dateEnd = lb.List(i, 2)
            For j = dateStart To dateEnd
                If j >= mdatecKorStart And j <= mdatecKorEnd Then
                    CalcKorea = True
                End If
            Next
        Next
End Function

Public Function CalcRVN(lb As MSForms.ListBox) As Boolean
        Dim i As Integer
        Dim j As Date
        Dim dateStart As Date
        Dim dateEnd As Date
        
        For i = 0 To lb.ListCount - 1
            dateStart = lb.List(i, 1)
            dateEnd = lb.List(i, 2)
            For j = dateStart To dateEnd
                If j >= mdatecRVNStart And j <= mdatecRVNEnd Then
                    CalcRVN = True
                End If
            Next
        Next
        
End Function

Public Function CalcGWOT(lb As MSForms.ListBox) As Boolean
        Dim i As Integer
        Dim j As Date
        Dim dateStart As Date
        Dim dateEnd As Date
        
        For i = 0 To lb.ListCount - 1
            dateStart = lb.List(i, 1)
            dateEnd = lb.List(i, 2)
            For j = dateStart To dateEnd
                If j >= mdatecGWOTStart Then
                    CalcGWOT = True
                End If
            Next
        Next
End Function

Public Function ClaimWithinYear(strEP As String, lb As MSForms.ListBox) As Boolean
    'This function will take an EP String and calculate whether the claim was filed
    'within a year of separation
    
    Dim intOpen As Integer
    Dim intClose As Integer
    Dim strDoc As String
    Dim dtDOC As Date
    Dim i As Integer
    Dim blnCheck As Boolean
    Dim intDur As Integer
    
    blnCheck = False
    
    intOpen = InStrRev(strEP, "(")
    intClose = InStrRev(strEP, ")")
    strDoc = Mid(strEP, intOpen + 1, intClose - intOpen - 1)
    dtDOC = CDate(strDoc)
    
    For i = 0 To lb.ListCount - 1
        intDur = modDateFunctions.DURATION(dtDOC, lb.List(i, 2))
        If intDur < 366 Then blnCheck = True
    Next
    
    ClaimWithinYear = blnCheck
    
End Function

Public Function DaysPending(strEP As String) As Integer
    'This function will take an EP String and Calculate Days Pending
    'by parsing what is between the parentheses and calculating
    'duration to today
    
    Dim intOpen As Integer
    Dim intClose As Integer
    Dim strDoc As String
    Dim dtDOC As Date

    intOpen = InStrRev(strEP, "(")
    intClose = InStrRev(strEP, ")")
    strDoc = Mid(strEP, intOpen + 1, intClose - intOpen - 1)
    dtDOC = CDate(strDoc)
    DaysPending = modDateFunctions.DURATION(dtDOC, Date)
End Function


Public Function ConsecNinety(lb As MSForms.ListBox) As Boolean

    Dim i As Integer
    Dim dateStart As Date
    Dim dateEnd As Date
    Dim intSvc As Integer
    Dim blnConsecutive As Boolean

    blnConsecutive = False
    intSvc = 0
    
    For i = 0 To lb.ListCount - 1
        dateStart = CDate(lb.List(i, 1))
        dateEnd = CDate(lb.List(i, 2))
        intSvc = modDateFunctions.DURATION(dateStart, dateEnd)
        If intSvc >= 90 Then
            blnConsecutive = True
        End If
        intSvc = 0
    Next

    ConsecNinety = blnConsecutive
    
End Function

Public Function ExactAge(BirthDate As Variant) As String
     
    Dim iYear As Integer
    Dim iMonth As Integer
    Dim d As Integer
    Dim dt As Date
    Dim sResult  As String
     
    If Not IsDate(BirthDate) Then Exit Function
     
    dt = CDate(BirthDate)
    If dt > now Then Exit Function
     
    iYear = Year(dt)
    iMonth = Month(dt)
    d = Day(dt)
    iYear = Year(Date) - iYear
    iMonth = Month(Date) - iMonth
    d = Day(Date) - d
     
    If Sgn(d) = -1 Then
        d = 30 - Abs(d)
        iMonth = iMonth - 1
    End If
     
    If Sgn(iMonth) = -1 Then
        iMonth = 12 - Abs(iMonth)
        iYear = iYear - 1
    End If
     
    sResult = iYear & "." & iMonth
     
    ExactAge = sResult
     
End Function


Public Function TotalService(lb As MSForms.ListBox) As Integer

    Dim i As Integer
    Dim dateStart As Date
    Dim dateEnd As Date
    Dim intSvc As Integer


    intSvc = 0

    For i = 0 To lb.ListCount - 1
        dateStart = CDate(lb.List(i, 1))
        dateEnd = CDate(lb.List(i, 2))
        intSvc = intSvc + modDateFunctions.DURATION(dateStart, dateEnd)
    Next

    TotalService = intSvc
    
End Function
Public Function VerifyService(lb As MSForms.ListBox) As Boolean
    'This is a function that will accept a listbox as a parameter and
    'return a true value if a period of service has missing dates or is unverified

    'Aragao Add
    Dim j As Integer
    
    For j = 0 To lb.ListCount - 1
            If (UCase(Trim(lb.List(j, 4))) = "NO" And UCase(Trim(lb.List(j, 5))) = "NO") Then
                VerifyService = True
            ElseIf (Trim(lb.List(j, 1) = "")) Then
                VerifyService = True
            ElseIf (Trim(lb.List(j, 2) = "")) Then
                VerifyService = True
            End If
    Next
End Function

Public Function DupService(lb As MSForms.ListBox) As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim dateStart As Date
    Dim dateEnd As Date
    Dim blnMatch As Boolean
    Dim intMatch As Integer
    
    blnMatch = False
    intMatch = 0
    
    For i = 0 To lb.ListCount - 1
        dateStart = CDate(lb.List(i, 1))
        dateEnd = CDate(lb.List(i, 2))
        
        For j = 0 To lb.ListCount - 1
            If dateStart = lb.List(j, 1) Or dateEnd = lb.List(j, 2) Then
                intMatch = intMatch + 1
            End If
            If intMatch > 1 Then blnMatch = True
        Next
        intMatch = 0
    Next
    
    DupService = blnMatch
    
End Function

'***************************************************************
'Controls/Form Info Not Passed as Parameters
'***************************************************************

Public Function DEATHPAYEE(dtDeath As Date) As Date
    DEATHPAYEE = WorksheetFunction.EoMonth((dtDeath), -1)
End Function

Public Function DRILLPAY(dtStart As Date, intDays As Integer) As Date
    
    Dim i As Integer
    Dim d As Date
    Dim d2 As Date
    
    i = 0
    d = dtStart
    
    Do Until i >= intDays
        
        i = WorksheetFunction.Days360(dtStart, d, True)
        If CStr(Left(d, 5) = "02/30") Then
            d = d + 2
        Else
            d = d + 1
        End If
    Loop

        d2 = d - 1
        
        If CStr(Left(d2, 5)) = "02/29" Then
            d2 = CDate("03/01/" & Right(d2, 4))
        End If
        
    DRILLPAY = d2
    
End Function

Public Function DDURATION(Date1 As Date, Date2 As Date)
    DDURATION = DateDiff("d", Date1, Date2) + 1
End Function

Public Function INCARCERATION(dtStart As Date) As Date
    INCARCERATION = DateAdd("d", 60, dtStart)
End Function

Public Function RESUMPTION(dtRAD As Date, dtClaim As Date) As Date
    Dim intDur As Integer
    Dim dtCalc As Date
    
    intDur = DURATION(dtRAD, dtClaim)
    
    If intDur <= 365 Then
        RESUMPTION = DateAdd("d", 1, dtRAD)
    Else
        dtCalc = DateAdd("yyyy", -1, dtClaim)
        If dtCalc = dtRAD Then
            dtCalc = DateAdd("d", 1, dtCalc)
        End If
        RESUMPTION = dtCalc
    End If
End Function

Public Function SIXTYPLUSFIRST(dtStart As Date) As Date
    SIXTYPLUSFIRST = WorksheetFunction.EoMonth((dtStart + 60), 0) + 1
End Function

Public Function SIXTYFIVEPLUSFIRST(dtStart As Date) As Date
    SIXTYFIVEPLUSFIRST = WorksheetFunction.EoMonth((dtStart + 65), 0) + 1
End Function

Public Function SIXTYPLUS(dtStart As Date) As Date
    SIXTYPLUS = DateAdd("d", 60, dtStart)
End Function

Public Function SIXTYPLUSLAST(dtStart As Date) As Date
    SIXTYPLUSLAST = WorksheetFunction.EoMonth((dtStart + 60), 0)
End Function



Public Function NPRC_STRS(strService As String, strComponent As String, dtRAD As Date) As Boolean
    
    NPRC_STRS = False
    
    Select Case strService
        Case "Army"
            If dtRAD < mdatecArmyNPRCS Then NPRC_STRS = True
        Case "Navy"
            If dtRAD < mdatecNavyNPRCS Then NPRC_STRS = True
        Case "Marine Corps"
            If dtRAD < mdatecUSMCNPRCS Then NPRC_STRS = True
        Case "Air Force"
            If strComponent = "Guard" Or strComponent = "Reserves" Then
                If dtRAD < mdatecAirForceNPRCSG Then NPRC_STRS = True
            ElseIf strComponent = "Active" Then
                If dtRAD < mdatecAirForceNPRCS Then NPRC_STRS = True
            End If
        Case "Coast Guard"
            If dtRAD < mdatecCoastNPRCS Then NPRC_STRS = True
    End Select
    
End Function

Public Function NPRC_OMPF(strService As String, dtRAD As Date) As Boolean
    
        NPRC_OMPF = False
        
        Select Case strService
            Case "Army"
                If dtRAD < mdatecArmyNPRCP Then NPRC_OMPF = True
            Case "Navy"
                If dtRAD < mdatecNavyNPRCP Then NPRC_OMPF = True
            Case "Marine Corps"
                If dtRAD < mdatecUSMCNPRCP Then NPRC_OMPF = True
            Case "Air Force"
                If dtRAD < mdatecAirForceNPRCP Then NPRC_OMPF = True
            Case "Coast Guard"
                NPRC_OMPF = True
        End Select
    
End Function

Public Function FIRERELATED(strService As String, dtRAD As Date) As Boolean
        
        FIRERELATED = False
        
        Select Case strService
            Case "Army"
                If dtRAD >= mdatecArmyFireSt And dtRAD <= mdatecArmyFireEnd Then FIRERELATED = True
            Case "Air Force"
                If dtRAD >= mdatecAirForceFireSt And dtRAD <= mdatecAirForceFireEnd Then FIRERELATED = True
            Case Else
                FIRERELATED = False
        End Select

End Function


Public Function HAIMS(strService As String, dtRAD As Date) As Boolean
    
    HAIMS = False
    
        Select Case strService
        
            Case "Army", "Navy", "Marine Corps", "Air Force"
                If dtRAD >= mdatecHAIMS Then HAIMS = True
            Case "Coast Guard"
                If dtRAD >= mdatecHAIMSCG Then HAIMS = True
        End Select

End Function

Public Function DPRIS(strService As String, dtRAD As Date) As Boolean
    
    DPRIS = False
    
    Select Case strService
        Case "Army"
            If dtRAD >= mdatecArmyDPRIS Then DPRIS = True
        Case "Navy"
            If dtRAD >= mdatecNavyDPRIS Then DPRIS = True
        Case "Air Force"
            If dtRAD >= mdatecAirForceDPRIS Then DPRIS = True
        Case "Marine Corps"
            If dtRAD >= mdatecMarinesDPRIS Then DPRIS = True
    End Select
    
End Function

Public Function RMC(strService As String, strComponent As String, dtRAD As Date) As Boolean

    RMC = False
    
    Select Case strService
        Case "Army"
            If dtRAD >= mdatecArmyNPRCS And dtRAD < mdatecHAIMS Then RMC = True
        Case "Navy"
            If dtRAD >= mdatecNavyNPRCS And dtRAD < mdatecHAIMS Then RMC = True
        Case "Air Force"
            If strComponent = "Guard" Or strComponent = "Reserves" Then
                If dtRAD >= mdatecAirForceNPRCSG And dtRAD < mdatecHAIMS Then RMC = True
            ElseIf strComponent = "Active" Then
                If dtRAD >= mdatecAirForceNPRCS And dtRAD < mdatecHAIMS Then RMC = True
            End If
        Case "Marine Corps"
            If dtRAD >= mdatecUSMCNPRCS And dtRAD < mdatecHAIMS Then RMC = True
        Case "Coast Guard"
            If dtRAD >= mdatecCoastNPRCS And dtRAD < mdatecHAIMSCG Then RMC = True
    End Select
End Function
