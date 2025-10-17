Attribute VB_Name = "ModEntryHelpers"
' Module: ModEntryHelpers (REPLACE the affected parts)
Option Explicit



Public Function BuildOutputText(practicalVal As String, reviewed As Boolean) As String
    ' New rules:
    ' - If reviewed checked:
    '     - If practical = "Incomplete" => "Reviewed"
    '     - Else => "Reviewed, " & <HarveyBall>
    ' - If not reviewed checked, return "" (invalid per your UI, but safe)
    If Not reviewed Then
        BuildOutputText = ""
        Exit Function
    End If
    
    Select Case practicalVal
        Case "Incomplete"
            BuildOutputText = "Reviewed"
        Case "0": BuildOutputText = "Reviewed, " & HB_Empty()   ' ?
        Case "1": BuildOutputText = "Reviewed, " & HB_Q1()      ' ?
        Case "2": BuildOutputText = "Reviewed, " & HB_Half()    ' ?
        Case "3": BuildOutputText = "Reviewed, " & HB_Q3()      ' ?
        Case "4": BuildOutputText = "Reviewed, " & HB_Full()    ' ?
        Case Else
            BuildOutputText = "Reviewed"
    End Select
End Function

' No more color coding of "Practical:" text (removed)
Public Sub ApplyPracticalColor(rng As Range, practicalVal As String)
    ' Deprecated by design change: clear to default font color
    rng.Font.Color = vbBlack
End Sub

' WriteValues also stamps per-cell dates on the very hidden date sheets
Public Sub WriteValues(ws As Worksheet, practicalVal As String, reviewed As Boolean, entryDate As Date)
    Dim i As Long, j As Long
    Dim opName As String, tisName As String
    Dim foundCol As Range, foundRow As Range
    Dim colNum As Long, rowNum As Long
    Dim cellTarget As Range
    Dim outputText As String
    
    With frmEntry
        For i = 0 To .lstOperator.ListCount - 1
            If .lstOperator.Selected(i) Then
                opName = .lstOperator.List(i)
                Set foundCol = ws.Rows(1).Find(What:=opName, LookAt:=xlWhole, MatchCase:=False)
                If foundCol Is Nothing Then GoTo NextOperator
                
                colNum = foundCol.Column
                
                For j = 0 To .lstTIS.ListCount - 1
                    If .lstTIS.Selected(j) Then
                        tisName = .lstTIS.List(j)
                        Set foundRow = ws.Columns(3).Find(What:=tisName, LookAt:=xlWhole, MatchCase:=False)
                        If foundRow Is Nothing Then GoTo NextTIS
                        
                        rowNum = foundRow.Row
                        Set cellTarget = ws.Cells(rowNum, colNum)
                        
                        outputText = BuildOutputText(practicalVal, reviewed)
                        cellTarget.Value = outputText
                        
                        ' apply color only to Practical portion (or reset for Incomplete)
                        ApplyPracticalColor cellTarget, practicalVal
                        
                        ' update hidden date sheets
                        If reviewed Then Call UpdateReviewDate(ws.Name, rowNum, colNum, entryDate)
                        If practicalVal <> "Incomplete" Then Call UpdatePracticalDate(ws.Name, rowNum, colNum, entryDate)
                    End If
NextTIS:
                Next j
            End If
NextOperator:
        Next i
    End With
End Sub


' Updated scoring: detect the ball directly (no "Practical:" marker)
Public Function GetScore(statusText As String) As Long
    Dim s As String: s = statusText
    s = Trim$(s)
    
    If s = "" Then
        GetScore = 0
        Exit Function
    End If
    
    If LCase$(Left$(s, 8)) = "reviewed" Then
        ' "Reviewed" or "Reviewed, <ball>"
        If InStr(1, s, HB_Full(), vbBinaryCompare) > 0 Then
            GetScore = 6
        ElseIf InStr(1, s, HB_Q3(), vbBinaryCompare) > 0 Then
            GetScore = 5
        ElseIf InStr(1, s, HB_Half(), vbBinaryCompare) > 0 Then
            GetScore = 4
        ElseIf InStr(1, s, HB_Q1(), vbBinaryCompare) > 0 Then
            GetScore = 3
        ElseIf InStr(1, s, HB_Empty(), vbBinaryCompare) > 0 Then
            GetScore = 2
        Else
            ' Reviewed only
            GetScore = 1
        End If
    ElseIf LCase$(Left$(s, 13)) = "update review" Then
        ' Treat as minimum credit (they need to re-review)
        GetScore = 1
    Else
        GetScore = 0
    End If
End Function

' === Write Review Date ===
Public Sub UpdateReviewDate(shiftName As String, rowNum As Long, colNum As Long, entryDate As Date)
    Dim wsDate As Worksheet
    On Error Resume Next
    Set wsDate = ThisWorkbook.Sheets(shiftName & " - Reviewed")
    On Error GoTo 0
    If wsDate Is Nothing Then Exit Sub
    
    wsDate.Cells(rowNum, colNum).Value = entryDate
End Sub

' === Write Practical Date ===
Public Sub UpdatePracticalDate(shiftName As String, rowNum As Long, colNum As Long, entryDate As Date)
    Dim wsDate As Worksheet
    On Error Resume Next
    Set wsDate = ThisWorkbook.Sheets(shiftName & " - Practical")
    On Error GoTo 0
    If wsDate Is Nothing Then Exit Sub
    
    wsDate.Cells(rowNum, colNum).Value = entryDate
End Sub

