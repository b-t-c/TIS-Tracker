VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConflicts 
   Caption         =   "Update Status Confirmation"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7410
   OleObjectBlob   =   "frmConflicts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConflicts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private storedConflicts As Collection
Private storedPractical As String
Private storedDate As Date    ' <<< new

' Pass conflicts from frmEntry
Public Sub SetConflicts(conflicts As Collection, practicalVal As String, entryDate As Date)
    Dim i As Long
    Dim item As Variant
    Set storedConflicts = conflicts
    storedPractical = practicalVal
    storedDate = entryDate     ' <<< capture the date
    
    lstConflicts.Clear
    For i = 1 To conflicts.Count
        item = conflicts(i)
        lstConflicts.AddItem item(0) & " | Current: " & item(1) & " | New: " & item(2)
        lstConflicts.AddItem " "
    Next i
End Sub


Private Sub btnConfirm_Click()
    ' overwrite values using the stored date
    WriteValues ActiveSheet, storedPractical, frmEntry.chkReviewed.Value, storedDate
    Unload Me
    Unload frmEntry
End Sub


Private Sub btnCancel_Click()
    ' just close this window and return to entry form
    Unload Me
End Sub

