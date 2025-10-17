Attribute VB_Name = "Module1"
Sub OpenEntryForm()
    frmEntry.Show
End Sub

Sub OpenTISManager()
    frmTISManager.Show
End Sub

Sub UnhideAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub

Sub HideSheetsExceptMain()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "White Days", "White Nights", "Orange Days", "Orange Nights", _
            "Summary, Operator %", "Summary, TIS vs. Shift %", "Summary, Full", _
            "TIS Master"
                ' keep visible
            Case Else
                ws.Visible = xlSheetHidden
        End Select
    Next ws
End Sub

Sub VeryHideSheetsExceptMain()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "White Days", "White Nights", "Orange Days", "Orange Nights", _
            "Summary, Operator %", "Summary, TIS vs. Shift %", "Summary, Full", _
            "TIS Master"
                ' keep visible
            Case Else
                ws.Visible = xlSheetVeryHidden
        End Select
    Next ws
End Sub

