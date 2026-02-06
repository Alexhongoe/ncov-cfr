Attribute VB_Name = "EASI_QC"
Option Explicit

Public Sub RunEasiQcSignals()
    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim lastRow As Long
    Dim headers As Object
    Dim rowIndex As Long
    Dim siteIndex As Long
    Dim siteCount As Long
    Dim siteIds() As String
    Dim siteNames() As String
    Dim baselineCount() As Long
    Dim baselineSum() As Double
    Dim baselineSumSq() As Double
    Dim avalCount() As Long
    Dim lastDigitCount() As Long
    Dim w16Count() As Long
    Dim w16OutlierCount() As Long
    Dim w16SiteIndex() As Long
    Dim w16Chg() As Double
    Dim w16Records As Long
    Dim globalChgSum As Double
    Dim globalChgSumSq As Double
    Dim globalChgN As Long
    Dim baselineSd() As Double
    Dim sdValues() As Double
    Dim sdCount As Long
    Dim baselineP10 As Double
    Dim baselineP05 As Double
    Dim colSiteId As Long
    Dim colSiteName As Long
    Dim colAvisit As Long
    Dim colAvisitn As Long
    Dim colAval As Long
    Dim colChg As Long

    Set ws = GetDataWorksheet()
    If ws Is Nothing Then
        MsgBox "Could not find a worksheet with data.", vbExclamation
        Exit Sub
    End If

    Set headers = GetHeaderIndex(ws)
    If Not headers.Exists("SITEID") Or Not headers.Exists("SITENAM") Or Not headers.Exists("AVISIT") Or _
       Not headers.Exists("AVISITN") Or Not headers.Exists("AVAL") Or Not headers.Exists("CHG") Then
        MsgBox "Missing required columns: SITEID, SITENAM, AVISIT, AVISITN, AVAL, CHG.", vbExclamation
        Exit Sub
    End If

    colSiteId = headers("SITEID")
    colSiteName = headers("SITENAM")
    colAvisit = headers("AVISIT")
    colAvisitn = headers("AVISITN")
    colAval = headers("AVAL")
    colChg = headers("CHG")

    lastRow = ws.Cells(ws.Rows.Count, colSiteId).End(xlUp).Row

    For rowIndex = 2 To lastRow
        Dim siteId As String
        Dim siteName As String
        Dim avisit As String
        Dim avisitn As Variant
        Dim aval As Variant
        Dim chg As Variant
        Dim intPart As Long
        Dim lastDigit As Long

        siteId = Trim(CStr(ws.Cells(rowIndex, colSiteId).Value))
        If Len(siteId) = 0 Then
            GoTo NextRow
        End If

        siteName = Trim(CStr(ws.Cells(rowIndex, colSiteName).Value))
        siteIndex = GetOrAddSite(siteId, siteName, siteIds, siteNames, baselineCount, baselineSum, baselineSumSq, _
                                 avalCount, lastDigitCount, w16Count, w16OutlierCount, siteCount)

        avisit = Trim(CStr(ws.Cells(rowIndex, colAvisit).Value))
        avisitn = ws.Cells(rowIndex, colAvisitn).Value
        aval = ws.Cells(rowIndex, colAval).Value
        chg = ws.Cells(rowIndex, colChg).Value

        If IsBaselineVisit(avisit, avisitn) Then
            If Not IsError(aval) And IsNumeric(aval) Then
                baselineCount(siteIndex) = baselineCount(siteIndex) + 1
                baselineSum(siteIndex) = baselineSum(siteIndex) + CDbl(aval)
                baselineSumSq(siteIndex) = baselineSumSq(siteIndex) + CDbl(aval) * CDbl(aval)
            End If
        End If

        If Not IsError(aval) And IsNumeric(aval) Then
            avalCount(siteIndex) = avalCount(siteIndex) + 1
            intPart = Fix(CDbl(aval))
            lastDigit = intPart Mod 10
            If lastDigit = 0 Or lastDigit = 5 Then
                lastDigitCount(siteIndex) = lastDigitCount(siteIndex) + 1
            End If
        End If

        If Not IsError(avisitn) And IsNumeric(avisitn) And CLng(avisitn) = 16 Then
            If Not IsError(chg) And IsNumeric(chg) Then
                w16Count(siteIndex) = w16Count(siteIndex) + 1
                w16Records = w16Records + 1
                ReDim Preserve w16SiteIndex(1 To w16Records)
                ReDim Preserve w16Chg(1 To w16Records)
                w16SiteIndex(w16Records) = siteIndex
                w16Chg(w16Records) = CDbl(chg)
                globalChgN = globalChgN + 1
                globalChgSum = globalChgSum + CDbl(chg)
                globalChgSumSq = globalChgSumSq + CDbl(chg) * CDbl(chg)
            End If
        End If
NextRow:
    Next rowIndex

    baselineSd = CalculateBaselineSd(siteCount, baselineCount, baselineSum, baselineSumSq)
    sdValues = ExtractSdValues(baselineSd, baselineCount, sdCount)
    If sdCount > 0 Then
        baselineP10 = Application.WorksheetFunction.Percentile_Inc(sdValues, 0.1)
        baselineP05 = Application.WorksheetFunction.Percentile_Inc(sdValues, 0.05)
    Else
        baselineP10 = 0
        baselineP05 = 0
    End If

    If w16Records > 1 Then
        Dim globalMean As Double
        Dim globalSd As Double
        globalMean = globalChgSum / globalChgN
        globalSd = SampleSd(globalChgN, globalChgSum, globalChgSumSq)
        If globalSd > 0 Then
            Dim recordIndex As Long
            For recordIndex = 1 To w16Records
                Dim diff As Double
                diff = Abs(w16Chg(recordIndex) - globalMean)
                If diff > 3 * globalSd Then
                    w16OutlierCount(w16SiteIndex(recordIndex)) = w16OutlierCount(w16SiteIndex(recordIndex)) + 1
                End If
            Next recordIndex
        End If
    End If

    Set outputWs = GetOrCreateOutputSheet("QC_SIGNALS")
    WriteOutput outputWs, siteCount, siteIds, siteNames, baselineCount, baselineSd, baselineP10, baselineP05, _
                avalCount, lastDigitCount, w16Count, w16OutlierCount

    MsgBox "QC signals written to sheet QC_SIGNALS.", vbInformation
End Sub

Private Function GetDataWorksheet() As Worksheet
    On Error Resume Next
    Set GetDataWorksheet = ThisWorkbook.Worksheets("DATA")
    On Error GoTo 0
    If GetDataWorksheet Is Nothing Then
        Set GetDataWorksheet = ThisWorkbook.ActiveSheet
    End If
End Function

Private Function GetHeaderIndex(ws As Worksheet) As Object
    Dim headers As Object
    Dim colIndex As Long
    Dim lastCol As Long
    Set headers = CreateObject("Scripting.Dictionary")

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For colIndex = 1 To lastCol
        Dim headerName As String
        headerName = Trim(UCase(CStr(ws.Cells(1, colIndex).Value)))
        If Len(headerName) > 0 Then
            headers(headerName) = colIndex
        End If
    Next colIndex

    Set GetHeaderIndex = headers
End Function

Private Function GetOrAddSite(siteId As String, siteName As String, _
                              ByRef siteIds() As String, ByRef siteNames() As String, _
                              ByRef baselineCount() As Long, ByRef baselineSum() As Double, ByRef baselineSumSq() As Double, _
                              ByRef avalCount() As Long, ByRef lastDigitCount() As Long, _
                              ByRef w16Count() As Long, ByRef w16OutlierCount() As Long, _
                              ByRef siteCount As Long) As Long
    Dim index As Long
    For index = 1 To siteCount
        If siteIds(index) = siteId Then
            GetOrAddSite = index
            Exit Function
        End If
    Next index

    siteCount = siteCount + 1
    ReDim Preserve siteIds(1 To siteCount)
    ReDim Preserve siteNames(1 To siteCount)
    ReDim Preserve baselineCount(1 To siteCount)
    ReDim Preserve baselineSum(1 To siteCount)
    ReDim Preserve baselineSumSq(1 To siteCount)
    ReDim Preserve avalCount(1 To siteCount)
    ReDim Preserve lastDigitCount(1 To siteCount)
    ReDim Preserve w16Count(1 To siteCount)
    ReDim Preserve w16OutlierCount(1 To siteCount)

    siteIds(siteCount) = siteId
    siteNames(siteCount) = siteName

    GetOrAddSite = siteCount
End Function

Private Function IsBaselineVisit(avisit As String, avisitn As Variant) As Boolean
    If LCase(avisit) = LCase("基线") Then
        IsBaselineVisit = True
        Exit Function
    End If
    If IsNumeric(avisitn) Then
        IsBaselineVisit = CLng(avisitn) = 0
    End If
End Function

Private Function CalculateBaselineSd(siteCount As Long, baselineCount() As Long, _
                                     baselineSum() As Double, baselineSumSq() As Double) As Double()
    Dim sdValues() As Double
    Dim index As Long
    ReDim sdValues(1 To siteCount)
    For index = 1 To siteCount
        sdValues(index) = SampleSd(baselineCount(index), baselineSum(index), baselineSumSq(index))
    Next index
    CalculateBaselineSd = sdValues
End Function

Private Function ExtractSdValues(baselineSd() As Double, baselineCount() As Long, ByRef sdCount As Long) As Double()
    Dim tmp() As Double
    Dim index As Long
    sdCount = 0
    For index = 1 To UBound(baselineSd)
        If baselineCount(index) > 1 Then
            sdCount = sdCount + 1
            ReDim Preserve tmp(1 To sdCount)
            tmp(sdCount) = baselineSd(index)
        End If
    Next index
    If sdCount = 0 Then
        ReDim tmp(1 To 1)
        tmp(1) = 0
        sdCount = 1
    End If
    ExtractSdValues = tmp
End Function

Private Function SampleSd(n As Long, sumValues As Double, sumSq As Double) As Double
    If n < 2 Then
        SampleSd = 0
    Else
        SampleSd = Sqr((sumSq - (sumValues * sumValues) / n) / (n - 1))
    End If
End Function

Private Function GetOrCreateOutputSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateOutputSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateOutputSheet Is Nothing Then
        Set GetOrCreateOutputSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateOutputSheet.Name = sheetName
    Else
        GetOrCreateOutputSheet.Cells.Clear
    End If
End Function

Private Sub WriteOutput(ws As Worksheet, siteCount As Long, siteIds() As String, siteNames() As String, _
                        baselineCount() As Long, baselineSd() As Double, baselineP10 As Double, baselineP05 As Double, _
                        avalCount() As Long, lastDigitCount() As Long, w16Count() As Long, w16OutlierCount() As Long)
    Dim rowIndex As Long
    ws.Range("A1:L1").Value = Array("SITEID", "SITENAM", "BaselineN", "BaselineSD", "BaselineFlag", _
                                    "LastDigitN", "LastDigit0or5Pct", "LastDigitFlag", _
                                    "W16N", "W16OutlierPct", "W16Flag", "Notes")
    For rowIndex = 1 To siteCount
        Dim baselineFlag As String
        Dim lastDigitFlag As String
        Dim w16Flag As String
        Dim lastDigitPct As Double
        Dim w16Pct As Double

        baselineFlag = "OK"
        If baselineSd(rowIndex) <= baselineP05 And baselineCount(rowIndex) >= 10 Then
            baselineFlag = "RED"
        ElseIf baselineSd(rowIndex) <= baselineP10 Then
            baselineFlag = "YELLOW"
        End If

        If avalCount(rowIndex) > 0 Then
            lastDigitPct = lastDigitCount(rowIndex) / avalCount(rowIndex)
        Else
            lastDigitPct = 0
        End If
        lastDigitFlag = "OK"
        If lastDigitPct > 0.6 And avalCount(rowIndex) >= 20 Then
            lastDigitFlag = "RED"
        ElseIf lastDigitPct > 0.45 Then
            lastDigitFlag = "YELLOW"
        End If

        If w16Count(rowIndex) > 0 Then
            w16Pct = w16OutlierCount(rowIndex) / w16Count(rowIndex)
        Else
            w16Pct = 0
        End If
        w16Flag = "OK"
        If w16Pct > 0.05 Then
            w16Flag = "RED"
        ElseIf w16Pct > 0.02 Then
            w16Flag = "YELLOW"
        End If

        ws.Cells(rowIndex + 1, 1).Value = siteIds(rowIndex)
        ws.Cells(rowIndex + 1, 2).Value = siteNames(rowIndex)
        ws.Cells(rowIndex + 1, 3).Value = baselineCount(rowIndex)
        ws.Cells(rowIndex + 1, 4).Value = Round(baselineSd(rowIndex), 4)
        ws.Cells(rowIndex + 1, 5).Value = baselineFlag
        ws.Cells(rowIndex + 1, 6).Value = avalCount(rowIndex)
        ws.Cells(rowIndex + 1, 7).Value = Round(lastDigitPct, 4)
        ws.Cells(rowIndex + 1, 8).Value = lastDigitFlag
        ws.Cells(rowIndex + 1, 9).Value = w16Count(rowIndex)
        ws.Cells(rowIndex + 1, 10).Value = Round(w16Pct, 4)
        ws.Cells(rowIndex + 1, 11).Value = w16Flag
        ws.Cells(rowIndex + 1, 12).Value = "Baseline SD P10=" & Round(baselineP10, 4) & "; P05=" & Round(baselineP05, 4)
    Next rowIndex

    ws.Columns.AutoFit
End Sub
