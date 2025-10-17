Attribute VB_Name = "ModTisSync"
' Module: ModTISSync
Option Explicit

' Ensure TIS Master, Archive exist; repair structure on shift sheets; push list & revisions; mark outdated upon revision deltas
Public Sub SyncTIS_All()
    EnsureMasterAndArchive
    EnsureDateSheets
    
    Dim arr, i As Long
    arr = ShiftSheets()
    
    Dim wsMaster As Worksheet: Set wsMaster = ThisWorkbook.Sheets(SHEET_TIS_MASTER)
    
    ' Read master list: DOC # (A), TIS Name (B), Revision (C)
    Dim lastRowM As Long: lastRowM = wsMaster.Cells(wsMaster.Rows.Count, 2).End(xlUp).Row
    If lastRowM < 2 Then Exit Sub
    
    ' For each shift, ensure columns/layout, sync list, push revisions, mark outdated
    For i = LBound(arr) To UBound(arr)
        SyncOneShift CStr(arr(i)), wsMaster, lastRowM
        ' Align date sheets grid
        SyncDateSheetToShift CStr(arr(i))
    Next i
    
    'MsgBox "TIS sync complete.", vbInformation
End Sub

Private Sub EnsureMasterAndArchive()
    Dim ws As Worksheet
    ' TIS Master already exists per your note
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_TIS_ARCHIVE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_TIS_ARCHIVE
        ws.Visible = xlSheetHidden
        ' Headers
        ws.Range("A1:I1").Value = Array("DOC #", "TIS Name", "RevisionAtDeletion", "Shift", "Operator", "CellText", "ReviewedDate", "PracticalDate", "DeletedOn")
    End If
End Sub

Private Sub EnsureShiftLayout(ByVal shiftName As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(shiftName)
    ' Make sure Column D = Revision exists (insert if needed),
    ' and operators start at G (COL_FIRST_OPERATOR)
    ' If you already inserted D in all shifts, this will generally be a no-op.
    
    ' If current col D isn't "Revision", and col C is TIS, we insert D
    If LCase$(Trim$(ws.Cells(1, COL_REV).Value)) <> "revision" Then
        ws.Columns(COL_REV).Insert
        ws.Cells(1, COL_REV).Value = "Revision"
    End If
    ' Ensure TIS header at C
    If LCase$(Trim$(ws.Cells(1, COL_TIS).Value)) <> "tis name" Then
        ws.Cells(1, COL_TIS).Value = "TIS Name"
    End If
End Sub

Private Sub SyncOneShift(ByVal shiftName As String, ByVal wsMaster As Worksheet, ByVal lastRowM As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(shiftName)
    Call EnsureShiftLayout(shiftName)
    
    Dim lastRowS As Long: lastRowS = ws.Cells(ws.Rows.Count, COL_TIS).End(xlUp).Row
    If lastRowS < 2 Then lastRowS = 1
    
    Dim r As Long, mName As String, mRev As Variant, mDoc As Variant
    Dim f As Range, rowNum As Long
    
    ' Build an index of existing TIS names on the shift sheet for quick find
    ' (we'll use .Find anyway to stay robust)
    
    For r = 2 To lastRowM
        mDoc = wsMaster.Cells(r, 1).Value
        mName = wsMaster.Cells(r, 2).Value
        mRev = wsMaster.Cells(r, 3).Value
        
        If Len(mName) > 0 Then
            Set f = ws.Columns(COL_TIS).Find(What:=mName, LookAt:=xlWhole, MatchCase:=False)
            If f Is Nothing Then
                ' Append new TIS row
                rowNum = ws.Cells(ws.Rows.Count, COL_TIS).End(xlUp).Row + 1
                ws.Cells(rowNum, COL_TIS - 1).Value = mDoc   ' DOC # in col B
                ws.Cells(rowNum, COL_TIS).Value = mName      ' TIS Name in col C
                ws.Cells(rowNum, COL_REV).Value = mRev       ' Revision in col D

            Else
                ' Exists -> check revision; if changed, mark "Update Review" on populated operator cells
                rowNum = f.Row
                If ws.Cells(rowNum, COL_REV).Value <> mRev Then
                    ws.Cells(rowNum, COL_TIS - 1).Value = mDoc   ' keep DOC # in sync too
                    ws.Cells(rowNum, COL_REV).Value = mRev
                    MarkRowOutdated shiftName, ws, rowNum
                End If
            End If
        End If
    Next r
End Sub

' For the specified TIS row on a shift sheet:
' - For each operator cell that's non-empty, replace "Reviewed" with "Update Review" (red),
'   preserve any existing Harvey ball after the comma
Private Sub MarkRowOutdated(ByVal shiftName As String, ByVal ws As Worksheet, ByVal rowNum As Long)
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, txt As String, commaPos As Long
    
    For c = COL_FIRST_OPERATOR To lastCol
        txt = CStr(ws.Cells(rowNum, c).Value)
        If Len(txt) > 0 Then
            ' Normalize leading token ? "Update Review"
            ' Cases:
            '   "Reviewed"                -> "Update Review"
            '   "Reviewed, <ball>"        -> "Update Review, <ball>"
            '   Already "Update Review..." -> leave as-is
            If Left$(LCase$(txt), 13) <> "update review" Then
                If LCase$(Left$(txt, 8)) = "reviewed" Then
                    ws.Cells(rowNum, c).Value = "Update Review" & Mid$(txt, 9) ' keep the rest (e.g., ", <ball>")
                Else
                    ' If a custom string existed, we still prefix with "Update Review, "
                    ' and try to keep any ball that may exist after comma
                    commaPos = InStr(1, txt, ",")
                    If commaPos > 0 Then
                        ws.Cells(rowNum, c).Value = "Update Review" & Mid$(txt, commaPos)
                    Else
                        ws.Cells(rowNum, c).Value = "Update Review"
                    End If
                End If
                ' Color only the "Update Review" phrase red
                With ws.Cells(rowNum, c)
                    .Characters(1, Len("Update Review")).Font.Color = RGB(192, 0, 0)
                End With
            End If
        End If
    Next c
End Sub

