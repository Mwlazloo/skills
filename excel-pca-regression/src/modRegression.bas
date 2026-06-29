Attribute VB_Name = "modRegression"
Option Explicit

' ============================================================
' OLS Regression Module
' ============================================================

' Run regression using PC scores from a saved PCA session.
' pcaSessionID  : the PCA session whose scores are X
' selectedPCIdx : 1-based array of which PCs to include  (e.g. Array(1,2,3))
' yRange        : worksheet range for the dependent variable (single column)
' addIntercept  : True to prepend a column of 1s
' sName         : friendly session name
' Returns       : session ID string
Function RunRegressionFromPCA(pcaSessionID As String, selectedPCIdx As Variant, _
                               yRange As Range, addIntercept As Boolean, sName As String) As String

    ' Build Y vector
    Dim yRaw As Variant: yRaw = yRange.Value
    Dim nObs As Long: nObs = UBound(yRaw, 1)
    Dim Y() As Double: ReDim Y(1 To nObs, 1 To 1)
    Dim i As Long
    For i = 1 To nObs
        If IsNumeric(yRaw(i, 1)) Then Y(i, 1) = CDbl(yRaw(i, 1))
    Next i

    ' Determine Y name
    Dim yName As String
    yName = yRange.Cells(1, 1).Address
    If yRange.Cells(1, 1).Row > 1 Then
        Dim above As String: above = yRange.Cells(1, 1).Offset(-1, 0).Value
        If above <> "" Then yName = above
    End If

    ' Get all PC scores then select the chosen columns
    Dim allScores As Variant: allScores = GetPCScores(pcaSessionID)
    Dim allPCNames As Variant: allPCNames = GetPCNames(pcaSessionID)
    Dim nSel As Long: nSel = UBound(selectedPCIdx) - LBound(selectedPCIdx) + 1
    Dim selScores() As Double: ReDim selScores(1 To nObs, 1 To nSel)
    Dim selNames() As String:  ReDim selNames(1 To nSel)
    Dim s As Long, pc As Long
    For s = 1 To nSel
        pc = selectedPCIdx(s - 1 + LBound(selectedPCIdx))
        selNames(s) = allPCNames(pc)
        Dim iObs As Long
        For iObs = 1 To nObs
            selScores(iObs, s) = allScores(iObs, pc)
        Next iObs
    Next s

    RunRegressionFromPCA = RunOLS(Y, selScores, yName, selNames, addIntercept, sName, pcaSessionID)
End Function

' Run regression using a custom worksheet range for X.
' xRange must be numeric data; xHasHeaders: True = first row is names
Function RunRegressionFromRange(xRange As Range, xHasHeaders As Boolean, _
                                 yRange As Range, yHasHeaders As Boolean, _
                                 addIntercept As Boolean, sName As String) As String

    Dim yRaw As Variant: yRaw = yRange.Value
    Dim xRaw As Variant: xRaw = xRange.Value
    Dim yStart As Long: If yHasHeaders Then yStart = 2 Else yStart = 1
    Dim xStart As Long: If xHasHeaders Then xStart = 2 Else xStart = 1

    Dim nObs As Long: nObs = UBound(yRaw, 1) - yStart + 1
    Dim nX As Long:   nX   = UBound(xRaw, 2)

    ' Y vector
    Dim Y() As Double: ReDim Y(1 To nObs, 1 To 1)
    Dim yName As String
    If yHasHeaders Then yName = CStr(yRaw(1, 1)) Else yName = "Y"
    Dim i As Long
    For i = 1 To nObs
        If IsNumeric(yRaw(i + yStart - 1, 1)) Then Y(i, 1) = CDbl(yRaw(i + yStart - 1, 1))
    Next i

    ' X matrix and names
    Dim xData() As Double: ReDim xData(1 To nObs, 1 To nX)
    Dim xNames() As String: ReDim xNames(1 To nX)
    Dim j As Long
    If xHasHeaders Then
        For j = 1 To nX: xNames(j) = CStr(xRaw(1, j)): Next j
    Else
        For j = 1 To nX: xNames(j) = "X" & j: Next j
    End If
    For i = 1 To nObs
        For j = 1 To nX
            If IsNumeric(xRaw(i + xStart - 1, j)) Then xData(i, j) = CDbl(xRaw(i + xStart - 1, j))
        Next j
    Next i

    RunRegressionFromRange = RunOLS(Y, xData, yName, xNames, addIntercept, sName, "")
End Function

' Core OLS: Y(n x 1), X(n x k), returns session ID
Private Function RunOLS(Y As Variant, Xmat As Variant, yName As String, xNames As Variant, _
                         addIntercept As Boolean, sName As String, pcaSID As String) As String

    Application.StatusBar = "Running OLS regression..."

    Dim nObs As Long: nObs = UBound(Y, 1)
    Dim nX As Long:   nX   = UBound(Xmat, 2)
    Dim k As Long:    k = nX + IIf(addIntercept, 1, 0)

    If nObs <= k Then
        MsgBox "Insufficient observations (n=" & nObs & ", k=" & k & "). Need n > k.", vbCritical
        Exit Function
    End If

    ' Build design matrix (optionally with intercept column)
    Dim X() As Double: ReDim X(1 To nObs, 1 To k)
    Dim i As Long, j As Long
    Dim colStart As Long: colStart = 1
    If addIntercept Then
        For i = 1 To nObs: X(i, 1) = 1#: Next i
        colStart = 2
    End If
    For i = 1 To nObs
        For j = 1 To nX
            X(i, j + colStart - 1) = Xmat(i, j)
        Next j
    Next i

    ' OLS:  beta = (X'X)^-1 X'Y
    Dim Xt As Variant:   Xt   = MatTranspose(X)
    Dim XtX As Variant:  XtX  = MatMul(Xt, X)
    Dim XtXi As Variant: XtXi = MatInverse(XtX)
    Dim XtY As Variant:  XtY  = MatMul(Xt, Y)
    Dim beta As Variant: beta = MatMul(XtXi, XtY)

    ' Fitted values and residuals
    Dim Yhat() As Double: ReDim Yhat(1 To nObs)
    Dim resid() As Double: ReDim resid(1 To nObs)
    Dim Ybar As Double: Ybar = 0#
    For i = 1 To nObs: Ybar = Ybar + Y(i, 1): Next i
    Ybar = Ybar / nObs

    Dim SSE As Double: SSE = 0#
    Dim SST As Double: SST = 0#
    Dim SSR As Double: SSR = 0#
    For i = 1 To nObs
        Dim yh As Double: yh = 0#
        For j = 1 To k: yh = yh + X(i, j) * beta(j, 1): Next j
        Yhat(i) = yh
        resid(i) = Y(i, 1) - yh
        SSE = SSE + resid(i) ^ 2
        SST = SST + (Y(i, 1) - Ybar) ^ 2
    Next i
    SSR = SST - SSE

    Dim df_err As Long: df_err = nObs - k
    Dim df_reg As Long: df_reg = k - IIf(addIntercept, 1, 0)
    If df_reg < 1 Then df_reg = 1
    Dim MSE As Double: If df_err > 0 Then MSE = SSE / df_err Else MSE = 0#
    Dim MSR As Double: If df_reg > 0 Then MSR = SSR / df_reg Else MSR = 0#
    Dim Rsq As Double:    If SST > 1E-15 Then Rsq = 1# - SSE / SST Else Rsq = 0#
    Dim AdjRsq As Double: If SST > 1E-15 And df_err > 0 Then AdjRsq = 1# - (SSE / df_err) / (SST / (nObs - 1)) Else AdjRsq = 0#
    Dim Fstat As Double:  If MSE > 1E-15 Then Fstat = MSR / MSE Else Fstat = 0#
    Dim Fpval As Double
    On Error Resume Next
    Fpval = WorksheetFunction.FDist(Fstat, df_reg, df_err)
    On Error GoTo 0

    ' Coefficient standard errors and t-stats
    Dim se() As Double:   ReDim se(1 To k)
    Dim tStat() As Double: ReDim tStat(1 To k)
    Dim pVal() As Double:  ReDim pVal(1 To k)
    For j = 1 To k
        se(j) = Sqr(MSE * XtXi(j, j))
        If se(j) > 1E-15 Then
            tStat(j) = beta(j, 1) / se(j)
        Else
            tStat(j) = 0#
        End If
        On Error Resume Next
        pVal(j) = WorksheetFunction.TDist(Abs(tStat(j)), df_err, 2)
        On Error GoTo 0
    Next j

    ' Build full variable name list (intercept + x names)
    Dim varList() As String: ReDim varList(1 To k)
    Dim v As Long: v = 1
    If addIntercept Then: varList(1) = "Intercept": v = 2
    For j = 1 To nX: varList(v + j - 1) = xNames(j): Next j

    ' Persist
    Dim sid As String: sid = GetNewID("REG")
    Call SaveREGData(sid, sName, yName, nObs, k, addIntercept, Rsq, AdjRsq, Fstat, Fpval, _
                     df_reg, df_err, varList, beta, se, tStat, pVal, resid, pcaSID)

    Application.StatusBar = False
    RunOLS = sid
End Function

' Read regression metadata for display
Function GetREGMeta(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim meta(1 To 9) As String
    meta(1) = CStr(ws.Cells(1, 2).Value)  ' name
    meta(2) = CStr(ws.Cells(1, 3).Value)  ' date
    meta(3) = CStr(ws.Cells(1, 4).Value)  ' Y name
    meta(4) = CStr(ws.Cells(2, 2).Value)  ' nObs
    meta(5) = CStr(ws.Cells(2, 3).Value)  ' k
    meta(6) = CStr(ws.Cells(3, 2).Value)  ' R2   (raw – form formats)
    meta(7) = CStr(ws.Cells(3, 3).Value)  ' AdjR2
    meta(8) = CStr(ws.Cells(3, 4).Value)  ' F
    meta(9) = CStr(ws.Cells(3, 5).Value)  ' F p-val
    GetREGMeta = meta
End Function

' Return coefficient table: rows = variables, cols = (name, coef, se, t, p, sig)
Function GetREGCoeffTable(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim k As Long: k = CLng(ws.Cells(2, 3).Value)
    Dim tbl() As String: ReDim tbl(1 To k, 1 To 6)
    Dim j As Long, pv As Double
    For j = 1 To k
        tbl(j, 1) = CStr(ws.Cells(5 + j, 1).Value)
        tbl(j, 2) = Format(CDbl(ws.Cells(5 + j, 2).Value), "0.00000")
        tbl(j, 3) = Format(CDbl(ws.Cells(5 + j, 3).Value), "0.00000")
        tbl(j, 4) = Format(CDbl(ws.Cells(5 + j, 4).Value), "0.000")
        pv = CDbl(ws.Cells(5 + j, 5).Value)
        tbl(j, 5) = Format(pv, "0.0000")
        If pv < 0.001 Then tbl(j, 6) = "***" _
        ElseIf pv < 0.01 Then tbl(j, 6) = "**" _
        ElseIf pv < 0.05 Then tbl(j, 6) = "*" _
        ElseIf pv < 0.1  Then tbl(j, 6) = "." _
        Else tbl(j, 6) = ""
    Next j
    GetREGCoeffTable = tbl
End Function

' Export full regression results to a new worksheet
Sub ExportRegressionToSheet(sessionID As String)
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Sub

    Dim outWS As Worksheet
    Dim outName As String: outName = "Reg_" & sessionID
    On Error Resume Next
    Set outWS = ThisWorkbook.Sheets(outName)
    On Error GoTo 0
    If outWS Is Nothing Then
        Set outWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        outWS.Name = outName
    Else
        outWS.Cells.ClearContents
    End If

    ' Copy all from session sheet
    ws.UsedRange.Copy outWS.Cells(1, 1)
    outWS.Activate
    MsgBox "Regression results copied to sheet '" & outName & "'.", vbInformation
End Sub
