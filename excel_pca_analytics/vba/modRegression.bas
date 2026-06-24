Attribute VB_Name = "modRegression"
'=============================================================================
' modRegression — OLS Regression and PCA Regression
'
' OLS formula:  beta = (X'X)^{-1} X'y
' Std errors:   se_j = sqrt(s² × [(X'X)^{-1}]_{jj})   where s² = SSR/(n-k-1)
' t-stats:      t_j  = beta_j / se_j
' p-values:     two-tailed t-distribution with (n-k-1) df
'=============================================================================
Option Explicit

' ── Standard OLS Regression ───────────────────────────────────────────────────
' X       : (1 To n, 1 To p) design matrix (WITHOUT intercept column)
' Y       : (1 To n) response vector
' incInt  : whether to prepend a column of 1s (intercept)
' Outputs written directly to REGRESSION sheet
Public Sub RunRegression()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Ensure PCA data has been loaded (need dates/Y/X from DATA)
    If g_n = 0 Then
        ' PCA hasn't been run yet — load data directly
        Dim Xtmp() As Double
        If Not ReadDataSheet(g_dates, g_Y, Xtmp, g_n, g_p) Then GoTo Cleanup
        g_scores = Xtmp   ' temporarily store raw X in g_scores for routing
        ' Actually we need original X, so re-read
        ReDim g_means(1 To g_p): ReDim g_stds(1 To g_p)
        Xtmp = Xtmp   ' already have it
        OLSRegress Xtmp, g_Y, g_n, g_p, "REGRESSION", False
    Else
        ' Re-read raw X (modPCA standardised its copy, we need original)
        Dim Xraw() As Double
        Dim dDates() As Variant: Dim dY() As Double
        Dim n2 As Long: Dim p2 As Long
        If Not ReadDataSheet(dDates, dY, Xraw, n2, p2) Then GoTo Cleanup
        OLSRegress Xraw, dY, n2, p2, "REGRESSION", False
    End If

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "OLS Regression complete.  Results on REGRESSION sheet.", _
           vbInformation, "PCA Analytics"
    Exit Sub
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ── PCA Regression ────────────────────────────────────────────────────────────
' Regresses Y on retained PC scores, then back-transforms to factor space.
Public Sub RunPCARegression()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Run PCA first if not done
    If g_n = 0 Then RunPCA
    If g_n = 0 Then GoTo Cleanup

    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Sheets("PCA_REG")
    Dim numPCs As Long: numPCs = CLng(wsC.Range("NUM_PCS").Value)

    ' Build include list from CONFIG Section C column E
    Dim inclPC() As Boolean: ReDim inclPC(1 To g_p)
    Dim pcCount As Long: pcCount = 0
    Dim j As Long
    For j = 1 To g_p
        Dim flag As String
        flag = UCase(Trim(wsC.Cells(27 + j, 5).Value))
        inclPC(j) = (flag = "YES")
        If inclPC(j) Then pcCount = pcCount + 1
    Next j
    If pcCount = 0 Then
        MsgBox "No PCs marked 'Yes' in CONFIG Section C column E.", vbExclamation
        GoTo Cleanup
    End If

    ' Build PC score matrix for included PCs only
    Dim Xpc() As Double
    ReDim Xpc(1 To g_n, 1 To pcCount)
    Dim colIdx As Long: colIdx = 0
    For j = 1 To g_p
        If inclPC(j) Then
            colIdx = colIdx + 1
            Dim i As Long
            For i = 1 To g_n
                Xpc(i, colIdx) = g_scores(i, j)
            Next i
        End If
    Next j

    ' Run OLS on PC scores
    OLSRegress Xpc, g_Y, g_n, pcCount, "PCA_REG", True

    ' ── Back-transform: factor contributions ──────────────────────────────────
    ' PC coeff vector (length pcCount); map back to p factors via loadings
    ' Contribution of factor i = sum_j (loading[i,j] * pcCoeff[j])  where j in retained
    Dim pcCoeffs() As Double: ReDim pcCoeffs(1 To pcCount)
    colIdx = 0
    For j = 1 To g_p
        If inclPC(j) Then
            colIdx = colIdx + 1
            ' VBA wrote coefficients to PCA_REG starting at row 15 (offset for intercept)
            pcCoeffs(colIdx) = wsR.Cells(15 + colIdx, 3).Value
        End If
    Next j

    Dim factorContrib() As Double: ReDim factorContrib(1 To g_p)
    Dim totalContribSq As Double: totalContribSq = 0
    For i = 1 To g_p
        Dim contrib As Double: contrib = 0
        colIdx = 0
        For j = 1 To g_p
            If inclPC(j) Then
                colIdx = colIdx + 1
                contrib = contrib + g_loadings(i, j) * pcCoeffs(colIdx)
            End If
        Next j
        factorContrib(i) = contrib
        totalContribSq = totalContribSq + contrib ^ 2
    Next i

    ' Write back-transformation
    For i = 1 To g_p
        wsR.Cells(28 + i, 3).Value = factorContrib(i)
        wsR.Cells(28 + i, 3).NumberFormat = "0.0000"
        If totalContribSq > 0 Then
            wsR.Cells(28 + i, 4).Value = factorContrib(i) ^ 2 / totalContribSq
            wsR.Cells(28 + i, 4).NumberFormat = "0.00%"
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "PCA Regression complete.  Results on PCA_REG sheet.", _
           vbInformation, "PCA Analytics"
    Exit Sub
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'=============================================================================
' Core OLS engine — writes results to the named worksheet
' sheetName : "REGRESSION" or "PCA_REG"
' isPCAReg  : affects label rows (PC labels vs Factor labels)
'=============================================================================
Private Sub OLSRegress(X() As Double, Y() As Double, n As Long, p As Long, _
                        sheetName As String, isPCAReg As Boolean)
    Dim ws  As Worksheet: Set ws = ThisWorkbook.Sheets(sheetName)
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")

    Dim incInt  As Boolean
    incInt = (UCase(Trim(wsC.Range("INCLUDE_INT").Value)) = "YES")
    Dim confLvl As Double
    confLvl = CDbl(wsC.Range("CONF_LEVEL").Value)
    Dim alpha   As Double: alpha = 1 - confLvl

    ' ── Build design matrix Xd (add intercept column if needed) ───────────────
    Dim k As Long: k = p + IIf(incInt, 1, 0)   ' total params
    Dim Xd() As Double: ReDim Xd(1 To n, 1 To k)
    Dim i As Long, j As Long
    If incInt Then
        For i = 1 To n: Xd(i, 1) = 1: Next i
        For j = 1 To p
            For i = 1 To n: Xd(i, j + 1) = X(i, j): Next i
        Next j
    Else
        For i = 1 To n
            For j = 1 To p: Xd(i, j) = X(i, j): Next j
        Next i
    End If

    ' ── X'X and X'y ───────────────────────────────────────────────────────────
    Dim XtX() As Double: XtX = MatMul(MatTrans(Xd), Xd)
    Dim Xty() As Double
    ReDim Xty(1 To k, 1 To 1)
    For j = 1 To k
        Dim s As Double: s = 0
        For i = 1 To n: s = s + Xd(i, j) * Y(i): Next i
        Xty(j, 1) = s
    Next j

    ' ── Invert X'X ────────────────────────────────────────────────────────────
    Dim XtXinv() As Double
    If Not MatInv(XtX, k, XtXinv) Then
        MsgBox "X'X is singular — multicollinearity detected. " & vbCrLf & _
               "Try reducing factors or using PCA Regression.", vbCritical
        Exit Sub
    End If

    ' ── Beta coefficients ──────────────────────────────────────────────────────
    Dim beta() As Double: beta = MatMul(XtXinv, Xty)

    ' ── Fitted values, residuals ───────────────────────────────────────────────
    Dim yHat() As Double: ReDim yHat(1 To n)
    Dim resid() As Double: ReDim resid(1 To n)
    Dim yMean As Double: yMean = 0
    For i = 1 To n: yMean = yMean + Y(i): Next i
    yMean = yMean / n

    Dim SSR As Double, SST As Double, SSE As Double
    For i = 1 To n
        yHat(i) = 0
        For j = 1 To k: yHat(i) = yHat(i) + Xd(i, j) * beta(j, 1): Next j
        resid(i) = Y(i) - yHat(i)
        SSE = SSE + resid(i) ^ 2
        SSR = SSR + (yHat(i) - yMean) ^ 2
        SST = SST + (Y(i) - yMean) ^ 2
    Next i

    ' ── Standard errors, t-stats, p-values ────────────────────────────────────
    Dim dfE As Long: dfE = n - k   ' residual degrees of freedom
    Dim s2  As Double: s2 = IIf(dfE > 0, SSE / dfE, 0)
    Dim tCrit As Double: tCrit = TInv(alpha, dfE)

    Dim seArr()   As Double: ReDim seArr(1 To k)
    Dim tArr()    As Double: ReDim tArr(1 To k)
    Dim pArr()    As Double: ReDim pArr(1 To k)
    Dim ciLo()    As Double: ReDim ciLo(1 To k)
    Dim ciHi()    As Double: ReDim ciHi(1 To k)
    Dim sigArr()  As String: ReDim sigArr(1 To k)

    For j = 1 To k
        seArr(j) = Sqr(IIf(s2 * XtXinv(j, j) >= 0, s2 * XtXinv(j, j), 0))
        tArr(j)  = IIf(seArr(j) > 0, beta(j, 1) / seArr(j), 0)
        pArr(j)  = TDistPValue(tArr(j), dfE)
        ciLo(j)  = beta(j, 1) - tCrit * seArr(j)
        ciHi(j)  = beta(j, 1) + tCrit * seArr(j)
        Select Case True
            Case pArr(j) < 0.01:  sigArr(j) = "***"
            Case pArr(j) < 0.05:  sigArr(j) = "**"
            Case pArr(j) < 0.1:   sigArr(j) = "*"
            Case Else:             sigArr(j) = ""
        End Select
    Next j

    ' ── Model summary stats ────────────────────────────────────────────────────
    Dim Rsq    As Double: Rsq    = IIf(SST > 0, 1 - SSE / SST, 0)
    Dim AdjRsq As Double: AdjRsq = IIf(dfE > 0, 1 - (1 - Rsq) * (n - 1) / dfE, 0)
    Dim dfR    As Long:   dfR    = k - IIf(incInt, 1, 0)
    Dim Fstat  As Double
    Dim pFstat As Double
    If dfR > 0 And s2 > 0 Then
        Fstat = (SSR / dfR) / s2
        ' F-dist p-value via beta function — approximate using chi-sq
        pFstat = FDistPValue(Fstat, dfR, dfE)
    End If
    Dim depName As String
    depName = CStr(wsC.Cells(4, 2).Value)

    ' ── AIC, BIC ──────────────────────────────────────────────────────────────
    Dim logLik  As Double
    If s2 > 0 Then
        logLik = -n / 2 * Log(2 * Application.WorksheetFunction.Pi() * s2) - SSE / (2 * s2)
    End If
    Dim AIC As Double: AIC = -2 * logLik + 2 * k
    Dim BIC As Double: BIC = -2 * logLik + Log(n) * k

    ' ══════════════════════════════════════════════════════════════════════════
    ' Write to worksheet
    ' ══════════════════════════════════════════════════════════════════════════
    Dim baseRow As Long
    If isPCAReg Then
        ' Model summary starts row 5
        ws.Cells(5, 2).Value = depName
        ws.Cells(6, 2).Value = pcCount(wsC)
        ws.Cells(7, 2).Value = n
        ws.Cells(8, 2).Value = Rsq:    ws.Cells(8, 2).NumberFormat = "0.0000"
        ws.Cells(9, 2).Value = AdjRsq: ws.Cells(9, 2).NumberFormat = "0.0000"
        ws.Cells(10, 2).Value = Fstat:  ws.Cells(10, 2).NumberFormat = "0.0000"
        ws.Cells(11, 2).Value = pFstat: ws.Cells(11, 2).NumberFormat = "0.0000"

        ' Coefficients (intercept then PCs)
        baseRow = 15
        If incInt Then
            ws.Cells(baseRow, 1).Value = "Intercept"
            ws.Cells(baseRow, 3).Value = beta(1, 1):  ws.Cells(baseRow, 3).NumberFormat = "0.0000"
            ws.Cells(baseRow, 4).Value = seArr(1):    ws.Cells(baseRow, 4).NumberFormat = "0.0000"
            ws.Cells(baseRow, 5).Value = tArr(1):     ws.Cells(baseRow, 5).NumberFormat = "0.0000"
            ws.Cells(baseRow, 6).Value = pArr(1):     ws.Cells(baseRow, 6).NumberFormat = "0.0000"
            ws.Cells(baseRow, 7).Value = ciLo(1):     ws.Cells(baseRow, 7).NumberFormat = "0.0000"
            ws.Cells(baseRow, 8).Value = ciHi(1):     ws.Cells(baseRow, 8).NumberFormat = "0.0000"
            baseRow = baseRow + 1
        End If
        ' PC rows
        Dim pcCol As Long: pcCol = 0
        Dim jBeta As Long: jBeta = IIf(incInt, 2, 1)
        For j = 1 To g_p
            Dim isIncl As Boolean
            isIncl = (UCase(Trim(wsC.Cells(27 + j, 5).Value)) = "YES")
            If isIncl Then
                ws.Cells(baseRow, 1).Value = "PC" & j
                ws.Cells(baseRow, 3).Value = beta(jBeta, 1): ws.Cells(baseRow, 3).NumberFormat = "0.0000"
                ws.Cells(baseRow, 4).Value = seArr(jBeta):   ws.Cells(baseRow, 4).NumberFormat = "0.0000"
                ws.Cells(baseRow, 5).Value = tArr(jBeta):    ws.Cells(baseRow, 5).NumberFormat = "0.0000"
                ws.Cells(baseRow, 6).Value = pArr(jBeta):    ws.Cells(baseRow, 6).NumberFormat = "0.0000"
                ws.Cells(baseRow, 7).Value = ciLo(jBeta):    ws.Cells(baseRow, 7).NumberFormat = "0.0000"
                ws.Cells(baseRow, 8).Value = ciHi(jBeta):    ws.Cells(baseRow, 8).NumberFormat = "0.0000"
                baseRow = baseRow + 1
                jBeta = jBeta + 1
            End If
        Next j
    Else
        ' Standard regression output (REGRESSION sheet)
        ' Model summary rows 5-12
        ws.Cells(5, 2).Value = depName
        ws.Cells(6, 2).Value = n
        ws.Cells(7, 2).Value = Rsq:    ws.Cells(7, 2).NumberFormat = "0.0000"
        ws.Cells(8, 2).Value = AdjRsq: ws.Cells(8, 2).NumberFormat = "0.0000"
        ws.Cells(9, 2).Value = Fstat:  ws.Cells(9, 2).NumberFormat = "0.0000"
        ws.Cells(10, 2).Value = pFstat: ws.Cells(10, 2).NumberFormat = "0.0000"
        ws.Cells(11, 2).Value = AIC:   ws.Cells(11, 2).NumberFormat = "0.00"
        ws.Cells(12, 2).Value = BIC:   ws.Cells(12, 2).NumberFormat = "0.00"

        ' ANOVA table rows 16-18
        ws.Cells(16, 2).Value = SSR: ws.Cells(16, 3).Value = dfR: ws.Cells(16, 4).Value = IIf(dfR > 0, SSR / dfR, 0)
        ws.Cells(16, 5).Value = Fstat: ws.Cells(16, 6).Value = pFstat
        ws.Cells(17, 2).Value = SSE: ws.Cells(17, 3).Value = dfE: ws.Cells(17, 4).Value = s2
        ws.Cells(18, 2).Value = SST: ws.Cells(18, 3).Value = n - 1
        For j = 16 To 18
            For i = 2 To 6: ws.Cells(j, i).NumberFormat = "0.0000": Next i
        Next j

        ' Coefficients rows 22+
        baseRow = 22
        Dim jOffset As Long: jOffset = 1
        If incInt Then
            WriteCoeffRow ws, baseRow, "Intercept", beta(1, 1), seArr(1), tArr(1), pArr(1), ciLo(1), ciHi(1), sigArr(1)
            baseRow = baseRow + 1
            jOffset = 2
        End If
        For j = 1 To p
            Dim factorName As String
            If isPCAReg Then
                factorName = "PC" & j
            Else
                factorName = CStr(wsC.Cells(14 + j, 3).Value)   ' CONFIG display name
                If Len(factorName) = 0 Then factorName = "Factor " & j
            End If
            WriteCoeffRow ws, baseRow, factorName, beta(jOffset + j - 1, 1), _
                          seArr(jOffset + j - 1), tArr(jOffset + j - 1), pArr(jOffset + j - 1), _
                          ciLo(jOffset + j - 1), ciHi(jOffset + j - 1), sigArr(jOffset + j - 1)
            baseRow = baseRow + 1
        Next j

        ' Residuals rows 37+
        For i = 1 To n
            ws.Cells(36 + i, 1).Value = g_dates(i)
            ws.Cells(36 + i, 1).NumberFormat = "YYYY-MM-DD"
            ws.Cells(36 + i, 2).Value = Y(i)
            ws.Cells(36 + i, 3).Value = yHat(i)
            ws.Cells(36 + i, 4).Value = resid(i)
            Dim stdResid As Double
            stdResid = IIf(Sqr(s2) > 0, resid(i) / Sqr(s2), 0)
            ws.Cells(36 + i, 5).Value = stdResid
            For j = 2 To 5: ws.Cells(36 + i, j).NumberFormat = "0.0000": Next j
        Next i
    End If
End Sub

' ── Write one coefficient row ──────────────────────────────────────────────────
Private Sub WriteCoeffRow(ws As Worksheet, r As Long, varName As String, _
                           b As Double, se As Double, t As Double, p As Double, _
                           ciL As Double, ciH As Double, sig As String)
    ws.Cells(r, 1).Value = varName
    ws.Cells(r, 2).Value = b:   ws.Cells(r, 2).NumberFormat = "0.0000"
    ws.Cells(r, 3).Value = se:  ws.Cells(r, 3).NumberFormat = "0.0000"
    ws.Cells(r, 4).Value = t:   ws.Cells(r, 4).NumberFormat = "0.0000"
    ws.Cells(r, 5).Value = p:   ws.Cells(r, 5).NumberFormat = "0.0000"
    ws.Cells(r, 6).Value = ciL: ws.Cells(r, 6).NumberFormat = "0.0000"
    ws.Cells(r, 7).Value = ciH: ws.Cells(r, 7).NumberFormat = "0.0000"
    ws.Cells(r, 8).Value = sig
    ' Color p-value
    With ws.Cells(r, 5).Font
        .Bold = True
        If p < 0.01 Then
            .Color = RGB(0, 128, 0)
        ElseIf p < 0.05 Then
            .Color = RGB(0, 176, 80)
        ElseIf p < 0.1 Then
            .Color = RGB(196, 139, 0)
        Else
            .Color = RGB(128, 128, 128)
        End If
    End With
End Sub

' ── Count retained PCs from CONFIG ────────────────────────────────────────────
Private Function pcCount(wsC As Worksheet) As Long
    Dim c As Long: c = 0
    Dim j As Long
    For j = 1 To 10
        If UCase(Trim(wsC.Cells(27 + j, 5).Value)) = "YES" Then c = c + 1
    Next j
    pcCount = c
End Function

' ── F-distribution p-value (via chi-squared approximation) ────────────────────
Private Function FDistPValue(F As Double, df1 As Long, df2 As Long) As Double
    On Error Resume Next
    FDistPValue = Application.WorksheetFunction.FDist(F, df1, df2)
    If Err.Number <> 0 Then FDistPValue = 1
    On Error GoTo 0
End Function
