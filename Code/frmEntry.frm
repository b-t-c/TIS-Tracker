VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntry 
   Caption         =   "Operator TIS Review/Training"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8370.001
   OleObjectBlob   =   "frmEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private allTIS As Collection


Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastCol As Long, lastRow As Long, i As Long
    
    Set ws = ActiveSheet ' <-- adjust if needed
    
    ' Populate Operator list (headers in row 1, starting F)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lstOperator.Clear
    For i = 7 To lastCol
        If ws.Cells(1, i).Value <> "" Then
            lstOperator.AddItem ws.Cells(1, i).Value
        End If
    Next i
    
    ' Populate TIS list (col C, starting row 2)
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    lstTIS.Clear
    Set allTIS = New Collection
    For i = 2 To lastRow
        If ws.Cells(i, 3).Value <> "" Then
            lstTIS.AddItem ws.Cells(i, 3).Value
            allTIS.Add ws.Cells(i, 3).Value
        End If
    Next i
    
    ' Populate Practical dropdown
    With cmbPractical
        .Clear
        .AddItem "Incomplete"
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .ListIndex = 0
    End With
End Sub

Private Sub txtSearchTIS_Change()
    Dim i As Long
    Dim searchText As String
    
    searchText = LCase(Trim(txtSearchTIS.Value))
    
    lstTIS.Clear
    For i = 1 To allTIS.Count
        If searchText = "" Or InStr(1, LCase(allTIS(i)), searchText) > 0 Then
            lstTIS.AddItem allTIS(i)
        End If
    Next i
End Sub

Private Sub cmbPractical_Change()
    If cmbPractical.Value <> "Incomplete" Then
        chkReviewed.Value = True
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim opName As String, tisName As String
    Dim colNum As Long, rowNum As Long
    Dim cellTarget As Range
    Dim practicalVal As String, outputText As String
    Dim conflicts As Collection
    Dim conflictInfo As Variant
    Dim entryDate As Date
    
    Set ws = ActiveSheet
    practicalVal = cmbPractical.Value
    
    ' --- determine entry date ---
    If Trim(txtDate.Value) = "" Then
        entryDate = Date ' default = today
    ElseIf IsDate(txtDate.Value) Then
        entryDate = CDate(txtDate.Value)
    Else
        MsgBox "Invalid date entered. Please use a valid date or leave blank to use today.", vbExclamation
        Exit Sub
    End If
    
    ' Validation
    If lstOperator.ListIndex = -1 Or lstTIS.ListIndex = -1 Or chkReviewed.Value = False Then
        MsgBox "Invalid Submission" & vbCrLf & _
               "At least one Operator, TIS, and an entry (Reviewed and/or Practical) must be selected.", _
               vbCritical, "Invalid Submission"
        Exit Sub
    End If
    
    ' Create a collection to store conflicts
    Set conflicts = New Collection
    
    ' Build list of targets
    For i = 0 To lstOperator.ListCount - 1
        If lstOperator.Selected(i) Then
            opName = lstOperator.List(i)
            colNum = ws.Rows(1).Find(What:=opName, LookAt:=xlWhole).Column
            
            For j = 0 To lstTIS.ListCount - 1
                If lstTIS.Selected(j) Then
                    tisName = lstTIS.List(j)
                    rowNum = ws.Columns(3).Find(What:=tisName, LookAt:=xlWhole).Row
                    Set cellTarget = ws.Cells(rowNum, colNum)
                    
                    ' Determine new value
                    outputText = BuildOutputText(practicalVal, chkReviewed.Value)
                    
                    ' Check for conflict
                    If cellTarget.Value <> "" Then
                        ' store: Address, CurrentValue, NewValue
                        conflictInfo = Array(cellTarget.Address, cellTarget.Value, outputText, entryDate)
                        conflicts.Add conflictInfo
                    End If
                End If
            Next j
        End If
    Next i
    
    ' If conflicts exist, launch conflict form
    If conflicts.Count > 0 Then
        frmConflicts.SetConflicts conflicts, practicalVal, entryDate
        frmConflicts.Show
    Else
        ' No conflicts: write immediately
        WriteValues ws, practicalVal, chkReviewed.Value, entryDate
        Unload Me
    End If
End Sub

