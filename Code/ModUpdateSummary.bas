Attribute VB_Name = "ModUpdateSummary"
Public Sub UpdateFullSummary()
    Dim wsFull As Worksheet, wsTIS As Worksheet, wsOps As Worksheet
    Dim ShiftSheets As Variant
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim nextRow As Long
    Dim statusText As String, practicalSymbol As String, reviewedOut As String
    Dim operatorName As Variant, tisName As Variant
    Dim dictTIS As Object, dictOperators As Object
    Dim totalScore As Long, maxScore As Long, score As Long
    Dim i As Long

    ' --- shift labels in the workbook ---
    ShiftSheets = Array("White Days", "White Nights", "Orange Days", "Orange Nights")
    
    ' --- create or clear target sheets ---
    On Error Resume Next
    Set wsFull = ThisWorkbook.Sheets("Summary, Full")
    Set wsTIS = ThisWorkbook.Sheets("Summary, TIS vs. Shift %")
    Set wsOps = ThisWorkbook.Sheets("Summary, Operator %")
    On Error GoTo 0
    
    If wsFull Is Nothing Then Set wsFull = ThisWorkbook.Sheets.Add: wsFull.Name = "Summary, Full" Else wsFull.Cells.Clear
    If wsTIS Is Nothing Then Set wsTIS = ThisWorkbook.Sheets.Add: wsTIS.Name = "Summary, TIS vs. Shift %" Else wsTIS.Cells.Clear
    If wsOps Is Nothing Then Set wsOps = ThisWorkbook.Sheets.Add: wsOps.Name = "Summary, Operator %" Else wsOps.Cells.Clear

    '==============================
    ' 1) Summary, Full (includes "No progress")
    '==============================
    wsFull.Range("A1:F1").Value = Array("Shift", "Operator", "TIS", "Status", "Practical Symbol", "Reviewed?")
    nextRow = 2

    For Each ws In ThisWorkbook.Sheets
        If Not IsError(Application.Match(Trim(ws.Name), ShiftSheets, 0)) Then
            lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row           ' TIS names in col C
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Operators from col F right

            For r = 2 To lastRow
                tisName = ws.Cells(r, 3).Value
                If Len(tisName) > 0 Then
                    For c = 7 To lastCol
                        operatorName = ws.Cells(1, c).Value
                        If Len(operatorName) > 0 Then
                            statusText = ws.Cells(r, c).Value

                            If Len(statusText) = 0 Then
                                ' Empty datapoint -> "No progress"
                                wsFull.Cells(nextRow, 1).Value = ws.Name
                                wsFull.Cells(nextRow, 2).Value = operatorName
                                wsFull.Cells(nextRow, 3).Value = tisName
                                wsFull.Cells(nextRow, 4).Value = "No progress"
                                wsFull.Cells(nextRow, 5).Value = ""   ' Practical Symbol
                                wsFull.Cells(nextRow, 6).Value = ""   ' Reviewed?
                            Else
                                ' Non-empty -> copy, extract symbol, reviewed?
                                practicalSymbol = ""
                                reviewedOut = ""
                                If InStr(statusText, "Practical:") > 0 Then
                                    ' one-char Harvey ball just after "Practical: "
                                    practicalSymbol = Mid(statusText, InStr(statusText, "Practical:") + 11, 1)
                                End If
                                If InStr(statusText, "Reviewed") > 0 Then reviewedOut = "Yes"

                                wsFull.Cells(nextRow, 1).Value = ws.Name
                                wsFull.Cells(nextRow, 2).Value = operatorName
                                wsFull.Cells(nextRow, 3).Value = tisName
                                wsFull.Cells(nextRow, 4).Value = statusText
                                wsFull.Cells(nextRow, 5).Value = practicalSymbol
                                wsFull.Cells(nextRow, 6).Value = reviewedOut
                            End If

                            nextRow = nextRow + 1
                        End If
                    Next c
                End If
            Next r
        End If
    Next ws

    ' Table + autofit
    Dim tblFull As ListObject
    On Error Resume Next
    Set tblFull = wsFull.ListObjects("tblSummaryFull")
    On Error GoTo 0
    
    ' unprotect sheet to allow table stuff
    wsFull.Unprotect Password:="1360"
    
    If tblFull Is Nothing Then
        Set tblFull = wsFull.ListObjects.Add(xlSrcRange, wsFull.Range("A1").CurrentRegion, , xlYes)
        tblFull.Name = "tblSummaryFull"
    Else
        tblFull.Resize wsFull.Range("A1").CurrentRegion
    End If
    wsFull.Columns.AutoFit

    ' --- color Shift cells in Summary, Full ---
    Dim lastRowFull As Long, cell As Range
    lastRowFull = wsFull.Cells(wsFull.Rows.Count, 1).End(xlUp).Row
    For Each cell In wsFull.Range("A2:A" & lastRowFull)
        Select Case cell.Value
            Case "White Days":   cell.Interior.Color = RGB(255, 255, 255)     ' White, Background 1
            Case "White Nights": cell.Interior.Color = RGB(191, 191, 191)     ' White, Background 1, Darker 25%
            Case "Orange Days":  cell.Interior.Color = RGB(255, 192, 0)       ' Orange, Accent 2
            Case "Orange Nights": cell.Interior.Color = RGB(192, 128, 0)      ' Orange, Accent 2, Darker 25%
        End Select
    Next cell
    
    ' reprotect sheet
    wsFull.Protect _
        Password:="1360", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        AllowFormattingCells:=False, _
        AllowSorting:=False, _
        AllowFiltering:=False, _
        AllowUsingPivotTables:=False, _
        UserInterfaceOnly:=True
    
    '==============================
    ' 2) Summary, TIS vs. Shift %
    '==============================
    Set dictTIS = CreateObject("Scripting.Dictionary")
    For Each ws In ThisWorkbook.Sheets
        If Not IsError(Application.Match(Trim(ws.Name), ShiftSheets, 0)) Then
            lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
            For r = 2 To lastRow
                tisName = ws.Cells(r, 3).Value
                If Len(tisName) > 0 Then
                    If Not dictTIS.exists(tisName) Then dictTIS.Add tisName, 1
                End If
            Next r
        End If
    Next ws

    wsTIS.Cells(1, 1).Value = "TIS"
    For i = 0 To UBound(ShiftSheets)
        wsTIS.Cells(1, 2 + i).Value = ShiftSheets(i) & " %"
    Next i

    nextRow = 2
    For Each tisName In dictTIS.Keys
        wsTIS.Cells(nextRow, 1).Value = tisName
        For i = 0 To UBound(ShiftSheets)
            Set ws = ThisWorkbook.Sheets(ShiftSheets(i))
            lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            Dim operatorCount As Long
            operatorCount = 0
            For c = 7 To lastCol
                If ws.Cells(1, c).Value <> "" Then operatorCount = operatorCount + 1
            Next c

            totalScore = 0
            maxScore = operatorCount * 6

            For r = 2 To lastRow
                If ws.Cells(r, 3).Value = tisName Then
                    For c = 7 To lastCol
                        If ws.Cells(1, c).Value <> "" Then
                            score = GetScore(ws.Cells(r, c).Value)
                            totalScore = totalScore + score
                        End If
                    Next c
                    Exit For
                End If
            Next r

            If maxScore > 0 Then
                wsTIS.Cells(nextRow, 2 + i).Value = Format(totalScore / maxScore, "0.0%")
            Else
                wsTIS.Cells(nextRow, 2 + i).Value = "N/A"
            End If
        Next i
        nextRow = nextRow + 1
    Next tisName

    Dim tblTISShift As ListObject
    On Error Resume Next
    Set tblTISShift = wsTIS.ListObjects("tblTISShift")
    On Error GoTo 0
    
    ' unprotect sheet to allow table stuff
    wsTIS.Unprotect Password:="1360"
    
    If tblTISShift Is Nothing Then
        Set tblTISShift = wsTIS.ListObjects.Add(xlSrcRange, wsTIS.Range("A1").CurrentRegion, , xlYes)
        tblTISShift.Name = "tblTISShift"
    Else
        tblTISShift.Resize wsTIS.Range("A1").CurrentRegion
    End If
    wsTIS.Columns.AutoFit

    ' --- color shift header cells in TIS vs Shift % ---
    Dim lastColTIS As Long
    lastColTIS = 1 + (UBound(ShiftSheets) + 1) ' col A is TIS, shifts start at col B
    For Each cell In wsTIS.Range(wsTIS.Cells(1, 1), wsTIS.Cells(1, lastColTIS))
        ' Always force header font to black for contrast
        cell.Font.Color = vbBlack
        Select Case Replace(cell.Value, " %", "")
            Case "White Days"
                cell.Interior.Color = RGB(255, 255, 255)
            Case "White Nights"
                cell.Interior.Color = RGB(191, 191, 191)
            Case "Orange Days"
                cell.Interior.Color = RGB(255, 192, 0)
            Case "Orange Nights"
                cell.Interior.Color = RGB(192, 128, 0)
        End Select
    Next cell
    
    ' reprotect sheet
    wsTIS.Protect _
        Password:="1360", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        AllowFormattingCells:=False, _
        AllowSorting:=False, _
        AllowFiltering:=False, _
        AllowUsingPivotTables:=False, _
        UserInterfaceOnly:=True
    
    '==============================
    ' 3) Summary, Operator % (expanded thresholds + most recent date)
    '==============================
    
    ' Build operator set (unique names) and map each operator to its sheet/column
    Set dictOperators = CreateObject("Scripting.Dictionary")
    
    For Each ws In ThisWorkbook.Sheets
        If Not IsError(Application.Match(Trim(ws.Name), ShiftSheets, 0)) Then
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            For c = 7 To lastCol
                operatorName = CStr(ws.Cells(1, c).Value)
                If Len(operatorName) > 0 Then
                    If Not dictOperators.exists(operatorName) Then
                        dictOperators.Add operatorName, ws.Name & "|" & c   ' store "shift|col"
                    End If
                End If
            Next c
        End If
    Next ws
    
    ' Headers row
    Dim headers As Variant
    headers = Array("Shift", "Operator", _
        "Reviewed", "Reviewed, " & HB_Empty(), "Reviewed, " & HB_Q1(), _
        "Reviewed, " & HB_Half(), "Reviewed, " & HB_Q3(), "Reviewed, " & HB_Full(), _
        "Most Recent Activity")
    wsOps.Cells(1, 1).Resize(1, UBound(headers) + 1).Value = headers
    
    nextRow = 2
    
    For Each operatorName In dictOperators.Keys
        Dim parts As Variant
        Dim opShift As String
        Dim opCol As Long
        Dim counts(1 To 6) As Long
        Dim totalTIS As Long
        Dim lastActivity As Date, checkDate As Variant
        
        ' Reset counters
        Erase counts
        totalTIS = 0
        lastActivity = 0
        
        ' Unpack shift|col
        parts = Split(dictOperators(operatorName), "|")
        opShift = parts(0)
        opCol = CLng(parts(1))
        
        ' Work within this operator's shift sheet only
        Set ws = ThisWorkbook.Sheets(opShift)
        lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
        
        For r = 2 To lastRow
            If Len(ws.Cells(r, 3).Value) > 0 Then
                totalTIS = totalTIS + 1
                score = GetScore(ws.Cells(r, opCol).Value)
                
                If score >= 1 Then counts(1) = counts(1) + 1
                If score >= 2 Then counts(2) = counts(2) + 1
                If score >= 3 Then counts(3) = counts(3) + 1
                If score >= 4 Then counts(4) = counts(4) + 1
                If score >= 5 Then counts(5) = counts(5) + 1
                If score >= 6 Then counts(6) = counts(6) + 1
                
                ' Track most recent review/practical dates
                checkDate = GetReviewedDate(ws.Name, r, opCol)
                If IsDate(checkDate) Then If checkDate > lastActivity Then lastActivity = checkDate
                checkDate = GetPracticalDate(ws.Name, r, opCol)
                If IsDate(checkDate) Then If checkDate > lastActivity Then lastActivity = checkDate
            End If
        Next r
        
        ' Output row
        wsOps.Cells(nextRow, 1).Value = opShift
        wsOps.Cells(nextRow, 2).Value = operatorName
        
        If totalTIS > 0 Then
            wsOps.Cells(nextRow, 3).Value = Format(counts(1) / totalTIS, "0.0%")
            wsOps.Cells(nextRow, 4).Value = Format(counts(2) / totalTIS, "0.0%")
            wsOps.Cells(nextRow, 5).Value = Format(counts(3) / totalTIS, "0.0%")
            wsOps.Cells(nextRow, 6).Value = Format(counts(4) / totalTIS, "0.0%")
            wsOps.Cells(nextRow, 7).Value = Format(counts(5) / totalTIS, "0.0%")
            wsOps.Cells(nextRow, 8).Value = Format(counts(6) / totalTIS, "0.0%")
        Else
            wsOps.Cells(nextRow, 3).Resize(1, 6).Value = "N/A"
        End If
        
        ' Most Recent Activity date
        If lastActivity > 0 Then
            wsOps.Cells(nextRow, 9).Value = lastActivity
            wsOps.Cells(nextRow, 9).NumberFormat = "mm/dd/yyyy"
            If lastActivity < Date - 14 Then
                wsOps.Cells(nextRow, 9).Font.Color = vbRed
            Else
                wsOps.Cells(nextRow, 9).Font.Color = vbBlack
            End If
        Else
            wsOps.Cells(nextRow, 9).Value = "N/A"
        End If
        
        nextRow = nextRow + 1
    Next operatorName
    
    ' Table + autofit for Operator %
    Dim tblOps As ListObject
    On Error Resume Next
    Set tblOps = wsOps.ListObjects("tblOperatorCompletion")
    On Error GoTo 0
    
    ' unprotect sheet to allow table stuff
    wsOps.Unprotect Password:="1360"
    
    If tblOps Is Nothing Then
        Set tblOps = wsOps.ListObjects.Add(xlSrcRange, wsOps.Range("A1").CurrentRegion, , xlYes)
        tblOps.Name = "tblOperatorCompletion"
    Else
        tblOps.Resize wsOps.Range("A1").CurrentRegion
    End If
    wsOps.Columns.AutoFit
    
    ' --- color Shift cells in Operator % ---
    Dim lastRowOps As Long, opCell As Range
    lastRowOps = wsOps.Cells(wsOps.Rows.Count, 1).End(xlUp).Row
    For Each opCell In wsOps.Range("A2:A" & lastRowOps)
        Select Case opCell.Value
            Case "White Days":   opCell.Interior.Color = RGB(255, 255, 255)
            Case "White Nights": opCell.Interior.Color = RGB(191, 191, 191)
            Case "Orange Days":  opCell.Interior.Color = RGB(255, 192, 0)
            Case "Orange Nights": opCell.Interior.Color = RGB(192, 128, 0)
        End Select
    Next opCell
    
    ' reprotect sheet
    wsOps.Protect _
        Password:="1360", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        AllowFormattingCells:=False, _
        AllowSorting:=False, _
        AllowFiltering:=False, _
        AllowUsingPivotTables:=False, _
        UserInterfaceOnly:=True
    
    Call CreateOperatorProgressChart
    'MsgBox "Full summary updated!", vbInformation
End Sub


