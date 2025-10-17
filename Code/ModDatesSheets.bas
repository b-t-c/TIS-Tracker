Attribute VB_Name = "ModDatesSheets"
' Module: ModDatesSheets
Option Explicit

' Ensure the two very hidden date sheets exist for each shift:
'   "<Shift> - Reviewed" and "<Shift> - Practical"
' Layout mirrors the shift sheet: row 1 has operator headers (G+),
' col C has TIS names. We'll keep the same structure so coordinates match.

Public Sub EnsureDateSheets()
    Dim arr, i As Long
    arr = ShiftSheets()
    For i = LBound(arr) To UBound(arr)
        EnsureOneDateSheet CStr(arr(i)), "Reviewed"
        EnsureOneDateSheet CStr(arr(i)), "Practical"
    Next i
End Sub

Private Sub EnsureOneDateSheet(ByVal shiftName As String, ByVal kind As String)
    Dim nm As String: nm = shiftName & " - " & kind
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nm
        'ws.Visible = xlSheetVeryHidden
    End If
End Sub

' Sync the row/column map of a single date sheet to its parent shift sheet:
' - Copy TIS names (col C)
' - Copy operator headers (row 1, G+)
Public Sub SyncDateSheetToShift(ByVal shiftName As String)
    Dim wsShift As Worksheet, wsRev As Worksheet, wsPrac As Worksheet
    Dim lastRow As Long, lastCol As Long
    
    Set wsShift = ThisWorkbook.Sheets(shiftName)
    Set wsRev = ThisWorkbook.Sheets(shiftName & " - Reviewed")
    Set wsPrac = ThisWorkbook.Sheets(shiftName & " - Practical")
    
    ' TIS names
    lastRow = wsShift.Cells(wsShift.Rows.Count, COL_TIS).End(xlUp).Row
    wsRev.Range("C:C").ClearContents
    wsPrac.Range("C:C").ClearContents
    If lastRow >= 2 Then
        wsRev.Range(wsRev.Cells(2, COL_TIS), wsRev.Cells(lastRow, COL_TIS)).Value = _
            wsShift.Range(wsShift.Cells(2, COL_TIS), wsShift.Cells(lastRow, COL_TIS)).Value
        wsPrac.Range(wsPrac.Cells(2, COL_TIS), wsPrac.Cells(lastRow, COL_TIS)).Value = _
            wsShift.Range(wsShift.Cells(2, COL_TIS), wsShift.Cells(lastRow, COL_TIS)).Value
    End If
    
    ' Operator headers
    lastCol = wsShift.Cells(1, wsShift.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_FIRST_OPERATOR Then Exit Sub
    
    wsRev.Rows(1).ClearContents
    wsPrac.Rows(1).ClearContents
    wsRev.Range(wsRev.Cells(1, COL_FIRST_OPERATOR), wsRev.Cells(1, lastCol)).Value = _
        wsShift.Range(wsShift.Cells(1, COL_FIRST_OPERATOR), wsShift.Cells(1, lastCol)).Value
    wsPrac.Range(wsPrac.Cells(1, COL_FIRST_OPERATOR), wsPrac.Cells(1, lastCol)).Value = _
        wsShift.Range(wsShift.Cells(1, COL_FIRST_OPERATOR), wsShift.Cells(1, lastCol)).Value
End Sub

' Helpers to write/read dates for a given TIS/op on a shift
Public Sub SetReviewedDate(ByVal shiftName As String, ByVal tisRow As Long, ByVal opCol As Long, ByVal dt As Date)
    With ThisWorkbook.Sheets(shiftName & " - Reviewed")
        .Cells(tisRow, opCol).Value = dt
        .Cells(tisRow, opCol).NumberFormat = "yyyy-mm-dd"
    End With
End Sub

Public Sub SetPracticalDate(ByVal shiftName As String, ByVal tisRow As Long, ByVal opCol As Long, ByVal dt As Date)
    With ThisWorkbook.Sheets(shiftName & " - Practical")
        .Cells(tisRow, opCol).Value = dt
        .Cells(tisRow, opCol).NumberFormat = "yyyy-mm-dd"
    End With
End Sub

Public Function GetReviewedDate(ByVal shiftName As String, ByVal tisRow As Long, ByVal opCol As Long) As Variant
    GetReviewedDate = ThisWorkbook.Sheets(shiftName & " - Reviewed").Cells(tisRow, opCol).Value
End Function

Public Function GetPracticalDate(ByVal shiftName As String, ByVal tisRow As Long, ByVal opCol As Long) As Variant
    GetPracticalDate = ThisWorkbook.Sheets(shiftName & " - Practical").Cells(tisRow, opCol).Value
End Function


