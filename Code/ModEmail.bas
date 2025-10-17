Attribute VB_Name = "ModEmail"


Sub DraftTrainingUpdateEmail()
    ' Always refresh summaries and chart first
    Call UpdateFullSummary
    Call CreateOperatorProgressChart
    
    Dim wsOps As Worksheet, tbl As ListObject
    Dim outApp As Object, outMail As Object
    Dim rngFiltered As Range, rngBody As Range
    Dim chartObj As ChartObject, tempPath As String, chartFile As String
    Dim i As Long, lastRow As Long
    
    Set wsOps = ThisWorkbook.Sheets("Summary, Operator %")
    On Error Resume Next
    Set tbl = wsOps.ListObjects("tblOperatorCompletion")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Operator % table not found. Run UpdateFullSummary first.", vbExclamation
        Exit Sub
    End If
    
    '--- Build filtered table range ---
    lastRow = tbl.DataBodyRange.Rows.Count
    Dim keepRows As Range
    
    For i = 1 To lastRow
        Dim val As Variant
        val = tbl.DataBodyRange.Cells(i, 9).Value ' Most Recent Activity col (I)
        If (IsError(val)) Or (val = "N/A") Or (Not IsDate(val)) Or (val < Date - 14) Then
            If keepRows Is Nothing Then
                Set keepRows = tbl.DataBodyRange.Rows(i)
            Else
                Set keepRows = Union(keepRows, tbl.DataBodyRange.Rows(i))
            End If
        End If
    Next i
    
    If keepRows Is Nothing Then
        MsgBox "No operators require follow-up. No email generated.", vbInformation
        Exit Sub
    End If
    
    '--- Build HTML table directly ---
    Dim HTMLbody As String
    Dim header As String
    Dim rowHTML As String
    Dim valHTML As Variant
    
    header = "<table border='1' cellspacing='0' cellpadding='5' style='border-collapse:collapse;font-family:Calibri;font-size:11pt;'>"
    header = header & "<tr style='background-color:#f2f2f2;font-weight:bold;'>"
    header = header & "<td>Shift</td><td>Operator</td><td>Most Recent Activity</td></tr>"
    
    HTMLbody = header
    
    For i = 1 To lastRow
        valHTML = tbl.DataBodyRange.Cells(i, 9).Value ' col I: Most Recent Activity
        If (IsError(valHTML)) Or (valHTML = "N/A") Or (Not IsDate(valHTML)) Or (valHTML < Date - 14) Then
            Dim shiftVal As String, operatorVal As String, activityVal As String
            shiftVal = tbl.DataBodyRange.Cells(i, 1).Value
            operatorVal = tbl.DataBodyRange.Cells(i, 2).Value
            activityVal = tbl.DataBodyRange.Cells(i, 9).Text
            
            ' default row style
            rowHTML = "<tr>"
            
            ' shift cell with background color
            Select Case shiftVal
                Case "White Days": rowHTML = rowHTML & "<td style='background-color:#FFFFFF;color:#000000;'>" & shiftVal & "</td>"
                Case "White Nights": rowHTML = rowHTML & "<td style='background-color:#BFBFBF;color:#000000;'>" & shiftVal & "</td>"
                Case "Orange Days": rowHTML = rowHTML & "<td style='background-color:#FFC000;color:#000000;'>" & shiftVal & "</td>"
                Case "Orange Nights": rowHTML = rowHTML & "<td style='background-color:#C08000;color:#000000;'>" & shiftVal & "</td>"
                Case Else: rowHTML = rowHTML & "<td>" & shiftVal & "</td>"
            End Select
            
            ' operator
            rowHTML = rowHTML & "<td>" & operatorVal & "</td>"
            
            ' activity, red font if overdue
            If (IsDate(valHTML) And valHTML < Date - 14) Or (valHTML = "N/A") Then
                rowHTML = rowHTML & "<td style='color:red;'>" & activityVal & "</td>"
            Else
                rowHTML = rowHTML & "<td>" & activityVal & "</td>"
            End If
            
            rowHTML = rowHTML & "</tr>"
            
            HTMLbody = HTMLbody & rowHTML
        End If
    Next i
    
    HTMLbody = HTMLbody & "</table>"

    
    '--- Export chart ---
    On Error Resume Next
    Set chartObj = wsOps.ChartObjects("OperatorProgressChart")
    On Error GoTo 0
    If chartObj Is Nothing Then
        MsgBox "Progress chart not found. Run CreateOperatorProgressChart first.", vbExclamation
        Exit Sub
    End If
    
    tempPath = Environ$("TEMP") & "\"
    chartFile = tempPath & "OperatorProgress.png"
    chartObj.Chart.Export Filename:=chartFile, FilterName:="PNG"
    
    '--- Create Outlook draft ---
    Dim wsRec As Worksheet, recRange As Range, recCell As Range
    Dim recList As String
    
    On Error Resume Next
    Set wsRec = ThisWorkbook.Sheets("UpdateRecipients")
    On Error GoTo 0
    
    If wsRec Is Nothing Then
        MsgBox "UpdateRecipients sheet not found. Please create it with recipient emails in column A.", vbExclamation
        Exit Sub
    End If
    
    ' Build recipient list from col A (non-empty values)
    Set recRange = wsRec.Range("A1", wsRec.Cells(wsRec.Rows.Count, 1).End(xlUp))
    recList = ""
    For Each recCell In recRange
        If Len(Trim(recCell.Value)) > 0 Then
            If recList = "" Then
                recList = recCell.Value
            Else
                recList = recList & ";" & recCell.Value
            End If
        End If
    Next recCell
    
    ' Create Outlook draft
    Set outApp = CreateObject("Outlook.Application")
    Set outMail = outApp.CreateItem(0)
    
    outMail.To = recList
    outMail.Subject = "TIS Review Update"
    outMail.HTMLbody = "<p>Facilities team,</p>" & _
                       "<p>Operators with >14 days since last review/assessment:</p>" & _
                       HTMLbody & _
                       "<p>Operator Harvey Ball status chart:</p>" & _
                       "<p><img src='" & chartFile & "'></p>" & _
                       "<p>Regards,<br>Automated TIS Training Tracker</p>"
    
    outMail.Display

    
    ' Attach chart to ensure it travels with the email
    outMail.Attachments.Add chartFile
End Sub



