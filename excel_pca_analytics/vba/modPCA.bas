Attribute VB_Name = "modPCA"
'=============================================================================
' modPCA — Principal Component Analysis
'
' Algorithm:
'   1. Read factor matrix X (n × p) from DATA sheet
'   2. Compute correlation or covariance matrix C (p × p)
'   3. Jacobi iterative eigendecomposition of C
'   4. Sort eigenvalues descending; fix sign convention
'   5. Compute factor loadings = eigenvectors × diag(sqrt(eigenvalues))
'   6. Compute PC scores = X_std × eigenvectors
'   7. Write all results to PCA sheet; hand off to modCharts
'=============================================================================
Option Explicit

' Public results — other modules read these after RunPCA completes
Public g_eigenVals()  As Double   ' (1 To p)
Public g_eigenVecs()  As Double   ' (1 To p, 1 To p) — columns = PCs
Public g_loadings()   As Double   ' (1 To p, 1 To p)
Public g_scores()     As Double   ' (1 To n, 1 To p)
Public g_covMat()     As Double   ' (1 To p, 1 To p)
Public g_means()      As Double   ' (1 To p)
Public g_stds()       As Double   ' (1 To p)
Public g_n            As Long
Public g_p            As Long
Public g_dates()      As Variant  ' (1 To n)
Public g_Y()          As Double   ' (1 To n) — dependent variable

' ── Main entry point ─────────────────────────────────────────────────────────
Public Sub RunPCA()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsP  As Worksheet: Set wsP  = ThisWorkbook.Sheets("PCA")
    Dim wsC  As Worksheet: Set wsC  = ThisWorkbook.Sheets("CONFIG")

    ' ── 1. Load data ──────────────────────────────────────────────────────────
    Dim X() As Double
    If Not ReadDataSheet(g_dates, g_Y, X, g_n, g_p) Then GoTo Cleanup

    Dim useCorrForm As Boolean
    useCorrForm = (UCase(Trim(wsC.Range("MATRIX_TYPE").Value)) = "CORRELATION")

    ' ── 2. Compute summary stats on original X (before standardising) ─────────
    ReDim g_means(1 To g_p)
    ReDim g_stds(1 To g_p)
    Dim mins() As Double:  ReDim mins(1 To g_p)
    Dim maxs() As Double:  ReDim maxs(1 To g_p)
    Dim j As Long, i As Long
    For j = 1 To g_p
        g_means(j) = ColMean(X, j)
        g_stds(j)  = ColStdDev(X, j, g_means(j))
        Dim mn As Double: mn = X(1, j)
        Dim mx As Double: mx = X(1, j)
        For i = 2 To g_n
            If X(i, j) < mn Then mn = X(i, j)
            If X(i, j) > mx Then mx = X(i, j)
        Next i
        mins(j) = mn: maxs(j) = mx
    Next j

    ' ── 3. Write Summary Statistics ───────────────────────────────────────────
    Dim r As Long
    For j = 1 To g_p
        r = 5 + j
        wsP.Cells(r, 2).Value = g_n
        wsP.Cells(r, 3).Value = g_means(j)
        wsP.Cells(r, 4).Value = g_stds(j)
        wsP.Cells(r, 5).Value = mins(j)
        wsP.Cells(r, 6).Value = maxs(j)
        wsP.Cells(r, 3).NumberFormat = "0.0000"
        wsP.Cells(r, 4).NumberFormat = "0.0000"
        wsP.Cells(r, 5).NumberFormat = "0.0000"
        wsP.Cells(r, 6).NumberFormat = "0.0000"
    Next j

    ' ── 4. Build correlation / covariance matrix ──────────────────────────────
    Dim Xstd() As Double
    ReDim Xstd(1 To g_n, 1 To g_p)
    For i = 1 To g_n
        For j = 1 To g_p: Xstd(i, j) = X(i, j): Next j
    Next i

    If useCorrForm Then
        ' Standardise Xstd (z-score) then compute sample covariance = correlation
        Dim dummy1() As Double: ReDim dummy1(1 To g_p)
        Dim dummy2() As Double: ReDim dummy2(1 To g_p)
        Standardize Xstd, g_n, g_p, dummy1, dummy2
    End If

    g_covMat = CovMatrix(Xstd, g_n, g_p, False)   ' already standardised above if corr

    ' Write correlation/cov matrix to sheet
    Dim matLabel As String
    matLabel = IIf(useCorrForm, "Correlation", "Covariance")
    wsP.Cells(17, 1).Value = "2.  " & UCase(matLabel) & " MATRIX"
    For i = 1 To g_p
        For j = 1 To g_p
            wsP.Cells(18 + i, j + 1).Value = g_covMat(i, j)
            wsP.Cells(18 + i, j + 1).NumberFormat = "0.0000"
        Next j
    Next i

    ' ── 5. Jacobi eigendecomposition ──────────────────────────────────────────
    ReDim g_eigenVals(1 To g_p)
    ReDim g_eigenVecs(1 To g_p, 1 To g_p)
    JacobiEigen g_covMat, g_p, g_eigenVals, g_eigenVecs

    ' Sort descending and fix sign
    SortEigenDesc g_eigenVals, g_eigenVecs, g_p
    FixEigenSigns g_eigenVecs, g_p

    ' ── 6. Factor Loadings = eigenvectors × sqrt(eigenvalues) ─────────────────
    ReDim g_loadings(1 To g_p, 1 To g_p)
    For i = 1 To g_p
        For j = 1 To g_p
            g_loadings(i, j) = g_eigenVecs(i, j) * Sqr(IIf(g_eigenVals(j) > 0, g_eigenVals(j), 0))
        Next j
    Next i

    ' ── 7. PC Scores = Xstd × eigenvectors ────────────────────────────────────
    g_scores = MatMul(Xstd, g_eigenVecs)

    ' ── 8. Write Eigenvalue Table ─────────────────────────────────────────────
    Dim totalVar As Double: totalVar = 0
    For j = 1 To g_p: totalVar = totalVar + g_eigenVals(j): Next j

    Dim cumPct As Double: cumPct = 0
    Dim numPCs As Long: numPCs = CLng(wsC.Range("NUM_PCS").Value)

    For j = 1 To g_p
        r = 31 + j
        Dim pctVar As Double
        pctVar = IIf(totalVar > 0, g_eigenVals(j) / totalVar, 0)
        cumPct = cumPct + pctVar
        wsP.Cells(r, 2).Value  = g_eigenVals(j)
        wsP.Cells(r, 3).Value  = pctVar
        wsP.Cells(r, 4).Value  = cumPct
        wsP.Cells(r, 5).Value  = IIf(j <= numPCs, "Retained", "Dropped")
        wsP.Cells(r, 2).NumberFormat = "0.0000"
        wsP.Cells(r, 3).NumberFormat = "0.00%"
        wsP.Cells(r, 4).NumberFormat = "0.00%"
        ' Color status
        With wsP.Cells(r, 5).Font
            .Bold = True
            .Color = IIf(j <= numPCs, RGB(0, 128, 0), RGB(180, 0, 0))
        End With
    Next j

    ' ── 9. Write Factor Loadings ───────────────────────────────────────────────
    For i = 1 To g_p
        For j = 1 To g_p
            wsP.Cells(44 + i, j + 1).Value = g_loadings(i, j)
            wsP.Cells(44 + i, j + 1).NumberFormat = "0.0000"
            ' Conditional color: positive=green, negative=red (soft)
            Dim lv As Double: lv = g_loadings(i, j)
            If Abs(lv) >= 0.5 Then
                wsP.Cells(44 + i, j + 1).Font.Bold = True
            Else
                wsP.Cells(44 + i, j + 1).Font.Bold = False
            End If
        Next j
    Next i

    ' ── 10. Write PC Scores ────────────────────────────────────────────────────
    For i = 1 To g_n
        wsP.Cells(57 + i, 1).Value = g_dates(i)
        wsP.Cells(57 + i, 1).NumberFormat = "YYYY-MM-DD"
        For j = 1 To numPCs
            wsP.Cells(57 + i, j + 1).Value = g_scores(i, j)
            wsP.Cells(57 + i, j + 1).NumberFormat = "0.0000"
        Next j
    Next i

    ' ── 11. Charts ────────────────────────────────────────────────────────────
    BuildPCACharts

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "PCA complete.  " & g_p & " factors, " & g_n & " observations." & vbCrLf & _
           "Results written to PCA sheet.", vbInformation, "PCA Analytics"
    Exit Sub
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

'=============================================================================
' Jacobi Iterative Eigendecomposition for real symmetric matrix
' A     : (1 To n, 1 To n) — input symmetric matrix (modified in place)
' n     : dimension
' eVals : (1 To n) — output eigenvalues
' eVecs : (1 To n, 1 To n) — output eigenvectors (columns)
'=============================================================================
Private Sub JacobiEigen(A() As Double, n As Long, _
                         eVals() As Double, eVecs() As Double)
    Dim i As Long, j As Long, k As Long
    Dim maxIter As Long: maxIter = 200 * n * n
    If maxIter < 500 Then maxIter = 500

    ' Init eigenvector matrix as identity
    ReDim eVecs(1 To n, 1 To n)
    For i = 1 To n: eVecs(i, i) = 1: Next i

    ' Work on a copy of A
    Dim B() As Double
    ReDim B(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n: B(i, j) = A(i, j): Next j
    Next i

    Dim p As Long, q As Long
    Dim maxOff As Double
    Dim theta As Double, c As Double, s As Double
    Dim bpp As Double, bqq As Double, bpq As Double
    Dim bip As Double, biq As Double
    Const TOL As Double = 1E-12

    Dim iter As Long
    For iter = 1 To maxIter
        ' Find off-diagonal element of largest absolute value
        maxOff = 0: p = 1: q = 2
        For i = 1 To n - 1
            For j = i + 1 To n
                If Abs(B(i, j)) > maxOff Then
                    maxOff = Abs(B(i, j))
                    p = i: q = j
                End If
            Next j
        Next i
        If maxOff < TOL Then Exit For   ' converged

        ' Compute rotation angle (using atan2 form for stability)
        Dim denom As Double
        denom = B(q, q) - B(p, p)
        If Abs(denom) < TOL Then
            theta = Application.WorksheetFunction.Pi() / 4
        Else
            theta = 0.5 * Atn(2 * B(p, q) / denom)
        End If
        c = Cos(theta)
        s = Sin(theta)

        ' Update diagonal and off-diagonal at (p,q)
        bpp = c * c * B(p, p) + 2 * s * c * B(p, q) + s * s * B(q, q)
        bqq = s * s * B(p, p) - 2 * s * c * B(p, q) + c * c * B(q, q)
        bpq = 0   ' zeroed by construction

        ' Update all other rows/cols k
        For k = 1 To n
            If k <> p And k <> q Then
                Dim bkp As Double: bkp =  c * B(k, p) + s * B(k, q)
                Dim bkq As Double: bkq = -s * B(k, p) + c * B(k, q)
                B(k, p) = bkp: B(p, k) = bkp
                B(k, q) = bkq: B(q, k) = bkq
            End If
        Next k
        B(p, p) = bpp: B(q, q) = bqq
        B(p, q) = bpq: B(q, p) = bpq

        ' Accumulate rotation in eigenvector columns p and q
        For i = 1 To n
            bip = eVecs(i, p)
            biq = eVecs(i, q)
            eVecs(i, p) =  c * bip + s * biq
            eVecs(i, q) = -s * bip + c * biq
        Next i
    Next iter

    ' Extract eigenvalues from diagonal
    ReDim eVals(1 To n)
    For i = 1 To n
        eVals(i) = B(i, i)
        If eVals(i) < 0 Then eVals(i) = 0   ' clamp floating-point negatives near zero
    Next i
End Sub
