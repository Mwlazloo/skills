Attribute VB_Name = "modMain"
'=============================================================================
' modMain — Entry points, keyboard shortcuts, and UI helpers
'
' Keyboard shortcuts registered in Auto_Open:
'   Ctrl+Shift+P  → RunPCA
'   Ctrl+Shift+R  → RunRegression
'   Ctrl+Shift+Q  → RunPCAReg
'   Ctrl+Shift+A  → RunAll
'   Ctrl+Shift+X  → ResetOutputs
'=============================================================================
Option Explicit

' ── Workbook open: register shortcuts ────────────────────────────────────────
Public Sub Auto_Open()
    Application.OnKey "^+P", "RunPCA"
    Application.OnKey "^+R", "RunRegression"
    Application.OnKey "^+Q", "RunPCAReg"
    Application.OnKey "^+A", "RunAll"
    Application.OnKey "^+X", "ResetOutputs"
End Sub

' ── Workbook close: deregister shortcuts ─────────────────────────────────────
Public Sub Auto_Close()
    Application.OnKey "^+P"
    Application.OnKey "^+R"
    Application.OnKey "^+Q"
    Application.OnKey "^+A"
    Application.OnKey "^+X"
End Sub

' ── Run PCA only ──────────────────────────────────────────────────────────────
Public Sub RunPCA_Entry()
    RunPCA
End Sub

' ── Run OLS Regression only ───────────────────────────────────────────────────
Public Sub RunRegression_Entry()
    RunRegression
End Sub

' ── Run PCA Regression only ───────────────────────────────────────────────────
Public Sub RunPCAReg_Entry()
    RunPCARegression
End Sub

' ── Run all three in sequence ─────────────────────────────────────────────────
Public Sub RunAll()
    Dim t0 As Single: t0 = Timer
    Application.ScreenUpdating = False

    RunPCA
    RunRegression
    RunPCARegression

    Application.ScreenUpdating = True
    Dim elapsed As Single: elapsed = Timer - t0
    MsgBox "All analyses complete in " & Format(elapsed, "0.0") & " seconds." & vbCrLf & vbCrLf & _
           "  PCA          → PCA sheet + CHARTS" & vbCrLf & _
           "  OLS Reg      → REGRESSION sheet" & vbCrLf & _
           "  PCA Reg      → PCA_REG sheet" & vbCrLf & vbCrLf & _
           "Rename PCs any time via CONFIG Section C.", _
           vbInformation, "PCA Analytics — Complete"
End Sub

' ── Clear all computed outputs (but not DATA or CONFIG) ───────────────────────
Public Sub ResetOutputs()
    Dim ans As Integer
    ans = MsgBox("Clear all computed outputs from PCA, REGRESSION, PCA_REG and CHARTS sheets?" & vbCrLf & _
                 "(DATA and CONFIG are not affected.)", vbYesNo + vbQuestion, "Reset Outputs")
    If ans <> vbYes Then Exit Sub

    Dim sheetNames As Variant
    sheetNames = Array("PCA", "REGRESSION", "PCA_REG", "CHARTS")

    Dim clearRanges As Variant
    ' Row ranges to clear computed values (headers stay because they're formulas linking to CONFIG)
    ' For PCA: stats rows 6:15, cov matrix rows 19:28, eigenvalues 32:41, loadings 45:54, scores 58:200
    ' For REGRESSION/PCA_REG: summary + tables
    Dim ws As Worksheet
    Dim nm As Variant

    Application.ScreenUpdating = False
    For Each nm In sheetNames
        Set ws = ThisWorkbook.Sheets(CStr(nm))
        Select Case CStr(nm)
            Case "PCA"
                ClearNumericRange ws, 6, 15, 2, 6     ' summary stats values
                ClearNumericRange ws, 19, 28, 2, 11   ' cov/corr matrix
                ClearNumericRange ws, 32, 41, 2, 6    ' eigenvalues
                ClearNumericRange ws, 45, 54, 2, 11   ' loadings
                ClearNumericRange ws, 58, 300, 1, 11  ' PC scores
                ' Clear temp chart data columns
                ClearNumericRange ws, 1, 10, 20, 50
                ClearNumericRange ws, 70, 90, 1, 15
            Case "REGRESSION"
                ClearNumericRange ws, 5, 12, 2, 2   ' model summary values
                ClearNumericRange ws, 16, 18, 2, 6  ' ANOVA
                ClearNumericRange ws, 22, 33, 2, 8  ' coefficients
                ClearNumericRange ws, 37, 300, 1, 5 ' residuals
            Case "PCA_REG"
                ClearNumericRange ws, 5, 11, 2, 2
                ClearNumericRange ws, 15, 26, 3, 8
                ClearNumericRange ws, 29, 38, 3, 4
            Case "CHARTS"
                Dim obj As ChartObject
                For Each obj In ws.ChartObjects: obj.Delete: Next obj
        End Select
    Next nm

    ' Reset global state
    g_n = 0: g_p = 0
    Application.ScreenUpdating = True
    MsgBox "Outputs cleared. Run analysis again when ready.", vbInformation, "PCA Analytics"
End Sub

' ── Navigate to a sheet ───────────────────────────────────────────────────────
Public Sub GoToData():       ThisWorkbook.Sheets("DATA").Activate: End Sub
Public Sub GoToConfig():     ThisWorkbook.Sheets("CONFIG").Activate: End Sub
Public Sub GoToPCA():        ThisWorkbook.Sheets("PCA").Activate: End Sub
Public Sub GoToRegression(): ThisWorkbook.Sheets("REGRESSION").Activate: End Sub
Public Sub GoToPCAReg():     ThisWorkbook.Sheets("PCA_REG").Activate: End Sub
Public Sub GoToCharts():     ThisWorkbook.Sheets("CHARTS").Activate: End Sub

' ── Utility: clear numeric values in a range (leave formulas intact) ──────────
Private Sub ClearNumericRange(ws As Worksheet, r1 As Long, r2 As Long, _
                               c1 As Long, c2 As Long)
    Dim r As Long, c As Long
    For r = r1 To r2
        For c = c1 To c2
            With ws.Cells(r, c)
                If Not .HasFormula Then .ClearContents
            End With
        Next c
    Next r
End Sub

' ── Refresh CONFIG factor names from DATA headers ─────────────────────────────
' Useful if the user rearranges DATA columns after initial setup.
Public Sub RefreshFactorNames()
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("DATA")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")
    Dim j As Long
    For j = 1 To 10
        Dim h As String: h = CStr(wsD.Cells(4, j + 2).Value)
        If Len(h) = 0 Then Exit For
        ' Only update CONFIG if the display name hasn't been customized
        Dim displayName As String: displayName = CStr(wsC.Cells(14 + j, 3).Value)
        Dim autoName As String:    autoName    = CStr(wsC.Cells(14 + j, 2).Value)
        If displayName = autoName Then
            wsC.Cells(14 + j, 3).Value = h   ' reset to new DATA header
        End If
    Next j
    MsgBox "Factor names refreshed from DATA sheet.", vbInformation, "PCA Analytics"
End Sub

' ── Quick diagnostic: show data dimensions and CONFIG summary ─────────────────
Public Sub ShowDiagnostics()
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("DATA")

    Dim startRow As Long: startRow = CLng(wsC.Range("DATA_START_ROW").Value)
    Dim endRow   As Long: endRow   = CLng(wsC.Range("DATA_END_ROW").Value)
    Dim nObs     As Long: nObs     = endRow - startRow + 1

    Dim p As Long: p = 0
    Dim j As Long
    For j = 3 To 30
        If wsD.Cells(4, j).Value = "" Then Exit For
        p = p + 1
    Next j

    Dim numPCs As Long: numPCs = CLng(wsC.Range("NUM_PCS").Value)
    Dim matType As String: matType = CStr(wsC.Range("MATRIX_TYPE").Value)
    Dim confLvl As Double: confLvl = CDbl(wsC.Range("CONF_LEVEL").Value)

    Dim msg As String
    msg = "PCA Analytics — Diagnostics" & vbCrLf & String(40, "─") & vbCrLf & vbCrLf
    msg = msg & "Data rows:         " & nObs & " (rows " & startRow & ":" & endRow & ")" & vbCrLf
    msg = msg & "Factors (p):       " & p & vbCrLf
    msg = msg & "PCs to retain:     " & numPCs & vbCrLf
    msg = msg & "Matrix type:       " & matType & vbCrLf
    msg = msg & "Confidence level:  " & Format(confLvl * 100, "0") & "%" & vbCrLf
    msg = msg & vbCrLf & "Global state:" & vbCrLf
    msg = msg & "  g_n (last run):  " & g_n & vbCrLf
    msg = msg & "  g_p (last run):  " & g_p & vbCrLf

    If g_n > 0 And g_p > 0 Then
        Dim totalVar As Double: Dim cumPct As Double
        For j = 1 To g_p: totalVar = totalVar + g_eigenVals(j): Next j
        msg = msg & vbCrLf & "Eigenvalue summary (top 5):" & vbCrLf
        Dim k As Long
        For k = 1 To IIf(g_p > 5, 5, g_p)
            Dim pct As Double
            pct = IIf(totalVar > 0, g_eigenVals(k) / totalVar, 0)
            cumPct = cumPct + pct
            msg = msg & "  PC" & k & ":  λ=" & Format(g_eigenVals(k), "0.0000") & _
                  "  (" & Format(pct * 100, "0.1") & "%)  cum=" & Format(cumPct * 100, "0.1") & "%" & vbCrLf
        Next k
    End If

    MsgBox msg, vbInformation, "Diagnostics"
End Sub

' ══════════════════════════════════════════════════════════════════════════════
' SETUP HELPER — run once to add buttons to each sheet
' Call from Immediate Window: modMain.AddSheetButtons
' ══════════════════════════════════════════════════════════════════════════════
Public Sub AddSheetButtons()
    Dim sheetBtns As Variant
    ' Each row: sheet name, button left, top, caption, macro name
    sheetBtns = Array( _
        Array("DATA",       10, 2, "▶ Run All (Ctrl+Shift+A)", "RunAll"), _
        Array("DATA",       200, 2, "⚙ Config", "GoToConfig"), _
        Array("CONFIG",     10, 2, "▶ Run All", "RunAll"), _
        Array("CONFIG",     160, 2, "↺ Refresh Factor Names", "RefreshFactorNames"), _
        Array("PCA",        10, 2, "▶ Run PCA (Ctrl+Shift+P)", "RunPCA_Entry"), _
        Array("PCA",        210, 2, "↻ Reset", "ResetOutputs"), _
        Array("REGRESSION", 10, 2, "▶ Run Regression (Ctrl+Shift+R)", "RunRegression_Entry"), _
        Array("PCA_REG",    10, 2, "▶ Run PCA Reg (Ctrl+Shift+Q)", "RunPCAReg_Entry"), _
        Array("CHARTS",     10, 2, "▶ Rebuild Charts", "RunAll") _
    )

    Dim item As Variant
    For Each item In sheetBtns
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CStr(item(0)))
        Dim btn As Shape
        Set btn = ws.Shapes.AddFormControl(xlButtonControl, _
            CSng(item(1)), CSng(item(2)), 170, 18)
        With btn
            .Name = "btn_" & CStr(item(4)) & "_" & ws.Name
            .TextFrame.Characters.Text = CStr(item(3))
            .TextFrame.Characters.Font.Size = 9
            .TextFrame.Characters.Font.Bold = True
            .OnAction = CStr(item(4))
            With .Fill
                .ForeColor.RGB = RGB(31, 56, 100)
                .BackColor.RGB = RGB(31, 56, 100)
            End With
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        End With
    Next item
    MsgBox "Buttons added to all sheets.", vbInformation, "Setup Complete"
End Sub
