Attribute VB_Name = "modMain"
Option Explicit

' ============================================================
' Main Entry Points
' ============================================================

Sub Auto_Open()
    Call InitStorage
End Sub

Sub ShowMainForm()
    Call InitStorage
    frmMain.Show
End Sub

' Quick-access button target (for any sheet button)
Sub OpenPCATool()
    Call ShowMainForm
End Sub

' Utility: format a Double as a percentage string
Function FmtPct(v As Double) As String
    FmtPct = Format(v * 100, "0.00") & "%"
End Function

' Utility: format p-value with significance stars
Function FmtPVal(p As Double) As String
    Dim stars As String
    If p < 0.001 Then stars = " ***" _
    ElseIf p < 0.01 Then stars = " **" _
    ElseIf p < 0.05 Then stars = " *" _
    ElseIf p < 0.1 Then stars = " ." _
    Else stars = ""
    FmtPVal = Format(p, "0.0000") & stars
End Function

' Utility: column-header letter from 1-based index
Function ColLetter(n As Long) As String
    ColLetter = Split(Cells(1, n).Address, "$")(1)
End Function

' Utility: truncate a string for display
Function Trunc(s As String, maxLen As Long) As String
    If Len(s) > maxLen Then Trunc = Left(s, maxLen - 1) & Chr(8230) Else Trunc = s
End Function
