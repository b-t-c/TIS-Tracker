VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddTIS 
   Caption         =   "Add New TIS"
   ClientHeight    =   2880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8820.001
   OleObjectBlob   =   "frmAddTIS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddTIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' UserForm: frmAddTIS
Option Explicit

Private Sub UserForm_Initialize()
    optNew.Value = True
    PopulateArchiveList
    ToggleMode
End Sub

Private Sub PopulateArchiveList()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_TIS_ARCHIVE)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    cmbReinstate.Clear
    Dim i As Long, seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        If Len(ws.Cells(i, 2).Value) > 0 Then
            Dim key As String: key = ws.Cells(i, 1).Value & "|" & ws.Cells(i, 2).Value
            If Not seen.exists(key) Then
                seen.Add key, 1
                cmbReinstate.AddItem ws.Cells(i, 1).Value & " | " & ws.Cells(i, 2).Value
            End If
        End If
    Next i
End Sub

Private Sub ToggleMode()
    Dim isNew As Boolean: isNew = optNew.Value
    txtDoc.Enabled = isNew
    txtName.Enabled = isNew
    txtRev.Enabled = isNew
    cmbReinstate.Enabled = Not isNew
End Sub

Private Sub optNew_Click(): ToggleMode: End Sub
Private Sub optReinstate_Click(): ToggleMode: End Sub

Private Sub btnOK_Click()
    If optNew.Value Then
        If Len(txtDoc.Value) = 0 Or Len(txtName.Value) = 0 Or Len(txtRev.Value) = 0 Then
            MsgBox "Please fill all fields", vbCritical
            Exit Sub
        End If
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_TIS_MASTER)
        Dim nextRow As Long: nextRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1
        ws.Cells(nextRow, 1).Value = txtDoc.Value
        ws.Cells(nextRow, 2).Value = txtName.Value
        ws.Cells(nextRow, 3).Value = txtRev.Value
        Call SyncTIS_All
    Else
        If cmbReinstate.ListIndex = -1 Then
            MsgBox "Select a TIS to reinstate.", vbCritical
            Exit Sub
        End If
        Dim parts() As String
        parts = Split(cmbReinstate.Value, " | ")
        Dim docNum As String: docNum = parts(0)
        Dim tisName As String: tisName = parts(1)
        Dim newRev As String
        newRev = InputBox("Enter revision for reinstated TIS " & tisName, "Reinstate TIS")
        If Len(newRev) = 0 Then Exit Sub
        Call ReinstateTIS(tisName, docNum, newRev)
    End If
    
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

