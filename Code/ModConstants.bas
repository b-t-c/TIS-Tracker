Attribute VB_Name = "ModConstants"
' Module: ModConstants
Option Explicit

Public Const SHEET_TIS_MASTER As String = "TIS Master"
Public Const SHEET_TIS_ARCHIVE As String = "TIS Archive"
' Structural columns (1-based)
Public Const COL_TIS As Long = 3       ' C: TIS Name
Public Const COL_REV As Long = 4       ' D: Revision
Public Const COL_FIRST_OPERATOR As Long = 7 ' G: first operator column

' Shift sheets (exact names)
Public Function ShiftSheets() As Variant
    ShiftSheets = Array("White Days", "White Nights", "Orange Days", "Orange Nights")
End Function


' Harvey balls (store by Unicode code points)
Public Function HB_Empty() As String: HB_Empty = ChrW(&H25CB) ' ?
End Function
Public Function HB_Q1() As String: HB_Q1 = ChrW(&H25D4) ' ?
End Function
Public Function HB_Half() As String: HB_Half = ChrW(&H25D1) ' ?
End Function
Public Function HB_Q3() As String: HB_Q3 = ChrW(&H25D5) ' ?
End Function
Public Function HB_Full() As String: HB_Full = ChrW(&H25CF) ' ?
End Function

