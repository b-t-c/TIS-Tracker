VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTISManager 
   Caption         =   "TIS List Modification"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5370
   OleObjectBlob   =   "frmTISManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTISManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' UserForm: frmTISManager
Option Explicit

Private Sub UserForm_Initialize()
    RefreshList
End Sub

Private Sub RefreshList()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SHEET_TIS_MASTER)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    lstTIS.Clear
    Dim i As Long
    For i = 2 To lastRow
        If Len(ws.Cells(i, 2).Value) > 0 Then
            lstTIS.AddItem ws.Cells(i, 1).Value & " | " & ws.Cells(i, 2).Value & " | " & ws.Cells(i, 3).Value
        End If
    Next i
End Sub

Private Sub btnAdd_Click()
    frmAddTIS.Show
    RefreshList
End Sub

Private Sub btnUpdate_Click()
    If lstTIS.ListIndex = -1 Then Exit Sub

    '--- Extract selected TIS info ---
    Dim parts() As String
    parts = Split(lstTIS.List(lstTIS.ListIndex), " | ")
    Dim oldDocNum As String: oldDocNum = parts(0)
    Dim oldTISName As String: oldTISName = parts(1)
    Dim oldRevDate As String: oldRevDate = parts(2)
    
    Dim wsMaster As Worksheet
    Set wsMaster = ThisWorkbook.Sheets(SHEET_TIS_MASTER)
    
    '--- Prompt for new DOC #, name, and revision ---
    Dim newDocNum As String, newName As String, newRev As String

    newDocNum = InputBox("Enter new DOC # (leave blank to keep '" & oldDocNum & "'):", _
                         "Update DOC #", oldDocNum)
    If Len(Trim(newDocNum)) = 0 Then newDocNum = oldDocNum

    newName = InputBox("Enter new TIS name (leave blank to keep '" & oldTISName & "'):", _
                       "Update TIS Name", oldTISName)
    If Len(Trim(newName)) = 0 Then newName = oldTISName

    newRev = InputBox("Enter new revision for '" & newName & "':", _
                      "Update Revision", oldRevDate)
    If Len(Trim(newRev)) = 0 Then newRev = oldRevDate
    
    '--- Update TIS Master ---
    Dim f As Range
    Set f = wsMaster.Columns(2).Find(What:=oldTISName, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        MsgBox "TIS not found on master sheet.", vbExclamation
        Exit Sub
    End If

    wsMaster.Cells(f.Row, 1).Value = newDocNum
    wsMaster.Cells(f.Row, 2).Value = newName
    wsMaster.Cells(f.Row, 3).Value = newRev
    
    '--- Update all related sheets ---
    Dim ws As Worksheet, r As Range
    Dim arrShifts As Variant
    arrShifts = Array("White Days", "White Nights", "Orange Days", "Orange Nights")

    For Each ws In ThisWorkbook.Sheets
        ' Include shift sheets and hidden review/practical sheets
        If Not IsError(Application.Match(ws.Name, arrShifts, 0)) _
           Or InStr(1, ws.Name, "ReviewedDates", vbTextCompare) > 0 _
           Or InStr(1, ws.Name, "PracticalDates", vbTextCompare) > 0 Then

            On Error Resume Next
            Set r = ws.Columns(COL_TIS).Find(What:=oldTISName, LookAt:=xlWhole, MatchCase:=False)
            On Error GoTo 0

            If Not r Is Nothing Then
                ws.Cells(r.Row, COL_TIS - 1).Value = newDocNum  ' update DOC #
                ws.Cells(r.Row, COL_TIS).Value = newName        ' update TIS name
                ws.Cells(r.Row, COL_REV).Value = newRev         ' update revision
            End If
        End If
    Next ws

    '--- Sync updates (handles marking outdated entries, etc.) ---
    Call SyncTIS_All

    '--- Refresh the list box ---
    RefreshList

    MsgBox "TIS updated successfully!" & vbCrLf & _
           "DOC #: " & newDocNum & vbCrLf & _
           "Name: " & newName & vbCrLf & _
           "Revision: " & newRev, vbInformation, "Update Complete"
End Sub



Private Sub btnDelete_Click()
    If lstTIS.ListIndex = -1 Then Exit Sub
    
    Dim parts() As String
    parts = Split(lstTIS.List(lstTIS.ListIndex), " | ")
    Dim tisName As String: tisName = parts(1)
    
    If MsgBox("Delete TIS '" & tisName & "'? This will archive data.", vbYesNo + vbExclamation) = vbYes Then
        Call ArchiveAndRemoveTIS(tisName)
    End If
    
    RefreshList
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

