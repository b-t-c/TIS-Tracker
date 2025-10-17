Attribute VB_Name = "ModTISArchive"

' Module: ModTISArchive
Option Explicit

' Capture an entire TIS row across all operators + dates into TIS Archive, then remove from shift sheets
Public Sub ArchiveAndRemoveTIS(ByVal tisName As String)
    Dim arr, i As Long, wsShift As Worksheet
    Dim lastCol As Long, lastRow As Long, r As Range, rowNum As Long
    Dim wsArch As Worksheet
    
    ' Ensure Archive sheet exists
    On Error Resume Next
    Set wsArch = ThisWorkbook.Sheets(SHEET_TIS_ARCHIVE)
    On Error GoTo 0
    
    If wsArch Is Nothing Then
        Set wsArch = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsArch.Name = SHEET_TIS_ARCHIVE
        wsArch.Visible = xlSheetHidden
        wsArch.Range("A1:I1").Value = Array("DOC #", "TIS Name", "RevisionAtDeletion", _
                                            "Shift", "Operator", "CellText", "ReviewedDate", _
                                            "PracticalDate", "DeletedOn")
    End If
    
    Dim docNum As Variant, revVal As Variant
    
    ' Pull DOC #, current Revision from master
    Dim wsM As Worksheet: Set wsM = ThisWorkbook.Sheets(SHEET_TIS_MASTER)
    Dim fM As Range: Set fM = wsM.Columns(2).Find(What:=tisName, LookAt:=xlWhole, MatchCase:=False)
    If fM Is Nothing Then Exit Sub
    docNum = wsM.Cells(fM.Row, 1).Value
    revVal = wsM.Cells(fM.Row, 3).Value
    
    arr = ShiftSheets()
    For i = LBound(arr) To UBound(arr)
        Set wsShift = ThisWorkbook.Sheets(CStr(arr(i)))
        Set r = wsShift.Columns(COL_TIS).Find(What:=tisName, LookAt:=xlWhole, MatchCase:=False)
        If Not r Is Nothing Then
            rowNum = r.Row
            lastCol = wsShift.Cells(1, wsShift.Columns.Count).End(xlToLeft).Column
            
            Dim c As Long, opName As String, cellText As String
            Dim revDate As Variant, pracDate As Variant
            For c = COL_FIRST_OPERATOR To lastCol
                opName = CStr(wsShift.Cells(1, c).Value)
                If Len(opName) > 0 Then
                    cellText = CStr(wsShift.Cells(rowNum, c).Value)
                    revDate = GetReviewedDate(wsShift.Name, rowNum, c)
                    pracDate = GetPracticalDate(wsShift.Name, rowNum, c)
                    
                    If Len(cellText) > 0 Or Not IsEmpty(revDate) Or Not IsEmpty(pracDate) Then
                        Dim nextRowA As Long
                        nextRowA = wsArch.Cells(wsArch.Rows.Count, 1).End(xlUp).Row + 1
                        wsArch.Cells(nextRowA, 1).Value = docNum
                        wsArch.Cells(nextRowA, 2).Value = tisName
                        wsArch.Cells(nextRowA, 3).Value = revVal
                        wsArch.Cells(nextRowA, 4).Value = wsShift.Name
                        wsArch.Cells(nextRowA, 5).Value = opName
                        wsArch.Cells(nextRowA, 6).Value = cellText
                        wsArch.Cells(nextRowA, 7).Value = revDate
                        wsArch.Cells(nextRowA, 8).Value = pracDate
                        wsArch.Cells(nextRowA, 9).Value = Date
                    End If
                End If
            Next c
            
            ' Remove the TIS row from this shift sheet
            If rowNum > 0 Then
                wsShift.Rows(rowNum).Delete
            End If
        End If
        
        ' Also keep date sheets aligned after deletion
        SyncDateSheetToShift wsShift.Name
    Next i
    
    ' Finally, remove from master
    wsM.Rows(fM.Row).Delete
End Sub



' Reinstate TIS across all shifts from the Archive (latest snapshot wins)
Public Sub ReinstateTIS(ByVal tisName As String, ByVal docNum As String, ByVal newRevision As Variant)
    Dim wsM As Worksheet: Set wsM = ThisWorkbook.Sheets(SHEET_TIS_MASTER)
    ' Append to master
    Dim nextM As Long: nextM = wsM.Cells(wsM.Rows.Count, 2).End(xlUp).Row + 1
    wsM.Cells(nextM, 1).Value = docNum
    wsM.Cells(nextM, 2).Value = tisName
    wsM.Cells(nextM, 3).Value = newRevision
    
    ' Push to shifts
    Call ModTisSync.SyncTIS_All
    
    ' Pull from archive and repopulate cells/dates where operators match
    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets(SHEET_TIS_ARCHIVE)
    Dim lastA As Long: lastA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    Dim i As Long, delRows As Collection
    Set delRows = New Collection
    
    For i = 2 To lastA
        If wsA.Cells(i, 2).Value = tisName Then
            ' restore operator data (existing code unchanged)
            Dim shiftName As String, op As String, txt As String
            Dim revDate As Variant, pracDate As Variant
            shiftName = wsA.Cells(i, 4).Value
            op = wsA.Cells(i, 5).Value
            txt = wsA.Cells(i, 6).Value
            revDate = wsA.Cells(i, 7).Value
            pracDate = wsA.Cells(i, 8).Value
            
            Dim wsS As Worksheet: Set wsS = ThisWorkbook.Sheets(shiftName)
            Dim fRow As Range, fCol As Range
            Set fRow = wsS.Columns(COL_TIS).Find(What:=tisName, LookAt:=xlWhole, MatchCase:=False)
            Set fCol = wsS.Rows(1).Find(What:=op, LookAt:=xlWhole, MatchCase:=False)
            If Not fRow Is Nothing And Not fCol Is Nothing Then
                Dim rowNum As Long, colNum As Long
                rowNum = fRow.Row
                colNum = fCol.Column
                
                wsS.Cells(rowNum, colNum).Value = NormalizeCellText(txt)
                If IsDate(revDate) Then SetReviewedDate shiftName, rowNum, colNum, revDate
                If IsDate(pracDate) Then SetPracticalDate shiftName, rowNum, colNum, pracDate
                
                ' revision check logic (same as before) ...
                Dim revVal As Variant
                revVal = wsM.Cells(nextM, 3).Value
                If IsDate(revDate) And IsDate(revVal) Then
                    If CDate(revDate) < CDate(revVal) Then
                        Dim txtNow As String: txtNow = wsS.Cells(rowNum, colNum).Value
                        If InStr(1, txtNow, ",") > 0 Then
                            wsS.Cells(rowNum, colNum).Value = "Update Review" & Mid$(txtNow, InStr(1, txtNow, ","))
                        Else
                            wsS.Cells(rowNum, colNum).Value = "Update Review"
                        End If
                        wsS.Cells(rowNum, colNum).Characters(1, Len("Update Review")).Font.Color = RGB(192, 0, 0)
                    End If
                End If
            End If
            
            ' Queue archive row for deletion
            delRows.Add i
        End If
    Next i
    
    ' --- Delete archived rows in reverse order ---
    Dim r As Variant
    For i = delRows.Count To 1 Step -1
        wsA.Rows(delRows(i)).Delete
    Next i
End Sub



Private Function NormalizeCellText(ByVal oldTxt As String) As String
    Dim t As String: t = Trim$(oldTxt)
    If t = "" Then NormalizeCellText = "": Exit Function
    
    ' Old formats may include "Reviewed, Practical: <ball>" or just "Reviewed"
    If InStr(1, t, "Practical:", vbTextCompare) > 0 Then
        ' Extract the ball (1 char after "Practical: ")
        Dim p As Long: p = InStr(1, t, "Practical:", vbTextCompare)
        Dim b As String: b = Mid$(t, p + 11, 1)
        NormalizeCellText = "Reviewed, " & b
    Else
        NormalizeCellText = t
    End If
End Function

