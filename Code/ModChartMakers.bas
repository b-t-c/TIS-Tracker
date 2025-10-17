Attribute VB_Name = "ModChartMakers"
Public Sub CreateOperatorProgressChart()
    Dim wsOps As Worksheet, tbl As ListObject
    Dim chartObj As ChartObject, cht As Chart
    Dim rowCount As Long, i As Long
    Dim opNames() As String
    Dim segData() As Double
    
    ' Locate Operator % sheet and table
    Set wsOps = ThisWorkbook.Sheets("Summary, Operator %")
    On Error Resume Next
    
    ' unprotect sheet to allow chart stuff
    wsOps.Unprotect Password:="1360"
    
    Set tbl = wsOps.ListObjects("tblOperatorCompletion")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Operator % table not found. Run UpdateFullSummary first.", vbExclamation
        Exit Sub
    End If
    
    ' Clear existing chart if present
    For Each chartObj In wsOps.ChartObjects
        If chartObj.Name = "OperatorProgressChart" Then chartObj.Delete
    Next chartObj
    
    rowCount = tbl.DataBodyRange.Rows.Count
    ReDim opNames(1 To rowCount)
    ReDim segData(1 To 6, 1 To rowCount) ' 6 stacked segments
    
    ' Build operator labels and segment data
    For i = 1 To rowCount
        opNames(i) = tbl.DataBodyRange.Cells(i, 1).Value & " - " & tbl.DataBodyRange.Cells(i, 2).Value
        
        Dim p1 As Double, p2 As Double, p3 As Double, p4 As Double, p5 As Double, p6 As Double
        p1 = tbl.DataBodyRange.Cells(i, 3).Value ' Reviewed
        p2 = tbl.DataBodyRange.Cells(i, 4).Value ' Reviewed+0
        p3 = tbl.DataBodyRange.Cells(i, 5).Value ' Reviewed+1
        p4 = tbl.DataBodyRange.Cells(i, 6).Value ' Reviewed+2
        p5 = tbl.DataBodyRange.Cells(i, 7).Value ' Reviewed+3
        p6 = tbl.DataBodyRange.Cells(i, 8).Value ' Reviewed+4
        
        ' Compute incremental segments (low to high)
        segData(1, i) = p1 - p2 ' Dark red
        segData(2, i) = p2 - p3 ' Crimson
        segData(3, i) = p3 - p4 ' Dark orange
        segData(4, i) = p4 - p5 ' Gold
        segData(5, i) = p5 - p6 ' Light green
        segData(6, i) = p6      ' Dark green
    Next i
    
    ' Create chart, auto-size height
    Dim chartHeight As Double
    chartHeight = WorksheetFunction.Min(800, WorksheetFunction.Max(200, rowCount * 18))
    
    Set chartObj = wsOps.ChartObjects.Add(Left:=wsOps.Cells(2, 11).Left, _
                                          Top:=wsOps.Cells(2, 11).Top, _
                                          Width:=600, Height:=chartHeight)
    chartObj.Name = "OperatorProgressChart"
    Set cht = chartObj.Chart
    cht.ChartType = xlBarStacked
    
    ' --- Add series in reverse order (dark green ? dark red) ---
    Dim s As Integer
    For s = 6 To 1 Step -1
        cht.SeriesCollection.NewSeries
        cht.SeriesCollection(cht.SeriesCollection.Count).Values = Application.Transpose(Application.Index(segData, s, 0))
        cht.SeriesCollection(cht.SeriesCollection.Count).XValues = opNames
        Select Case s
            Case 1: cht.SeriesCollection(cht.SeriesCollection.Count).Name = "Reviewed"
                    cht.SeriesCollection(cht.SeriesCollection.Count).Format.Fill.ForeColor.RGB = RGB(139, 0, 0) ' Dark red
            Case 2: cht.SeriesCollection(cht.SeriesCollection.Count).Name = "Reviewed, " & HB_Empty()
                    cht.SeriesCollection(cht.SeriesCollection.Count).Format.Fill.ForeColor.RGB = RGB(220, 20, 60) ' Crimson
            Case 3: cht.SeriesCollection(cht.SeriesCollection.Count).Name = "Reviewed, " & HB_Q1()
                    cht.SeriesCollection(cht.SeriesCollection.Count).Format.Fill.ForeColor.RGB = RGB(255, 140, 0) ' Dark orange
            Case 4: cht.SeriesCollection(cht.SeriesCollection.Count).Name = "Reviewed, " & HB_Half()
                    cht.SeriesCollection(cht.SeriesCollection.Count).Format.Fill.ForeColor.RGB = RGB(255, 215, 0) ' Gold
            Case 5: cht.SeriesCollection(cht.SeriesCollection.Count).Name = "Reviewed, " & HB_Q3()
                    cht.SeriesCollection(cht.SeriesCollection.Count).Format.Fill.ForeColor.RGB = RGB(144, 238, 144) ' Light green
            Case 6: cht.SeriesCollection(cht.SeriesCollection.Count).Name = "Reviewed, " & HB_Full()
                    cht.SeriesCollection(cht.SeriesCollection.Count).Format.Fill.ForeColor.RGB = RGB(0, 180, 20)   ' Dark green
        End Select
    Next s
    
    ' Axis formatting
    cht.Axes(xlValue).MaximumScale = 1
    cht.Axes(xlValue).MinimumScale = 0
    cht.Axes(xlValue).TickLabels.NumberFormat = "0%"
    cht.Axes(xlCategory).ReversePlotOrder = True
    
    cht.HasTitle = True
    cht.ChartTitle.Text = "Operator TIS Review/Assessment Status"
    cht.Legend.Position = xlBottom

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

End Sub


