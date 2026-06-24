Attribute VB_Name = "modUtils"
'=============================================================================
' modUtils — Matrix arithmetic and statistical helpers
' All arrays are 1-indexed (1 To n) throughout this workbook.
'=============================================================================
Option Explicit

' ── Matrix Multiply  C = A * B ───────────────────────────────────────────────
' A is (r x m), B is (m x c)
Public Function MatMul(A() As Double, B() As Double) As Double()
    Dim r As Long, m As Long, c As Long
    r = UBound(A, 1): m = UBound(A, 2): c = UBound(B, 2)
    Dim C() As Double
    ReDim C(1 To r, 1 To c)
    Dim i As Long, j As Long, k As Long, s As Double
    For i = 1 To r
        For j = 1 To c
            s = 0
            For k = 1 To m
                s = s + A(i, k) * B(k, j)
            Next k
            C(i, j) = s
        Next j
    Next i
    MatMul = C
End Function

' ── Transpose ─────────────────────────────────────────────────────────────────
Public Function MatTrans(A() As Double) As Double()
    Dim r As Long, c As Long
    r = UBound(A, 1): c = UBound(A, 2)
    Dim T() As Double
    ReDim T(1 To c, 1 To r)
    Dim i As Long, j As Long
    For i = 1 To r
        For j = 1 To c
            T(j, i) = A(i, j)
        Next j
    Next i
    MatTrans = T
End Function

' ── Inverse via Gauss-Jordan (in-place on augmented matrix) ──────────────────
' Returns False if matrix is singular (det ≈ 0).
Public Function MatInv(A() As Double, n As Long, Inv() As Double) As Boolean
    Dim aug() As Double
    ReDim aug(1 To n, 1 To 2 * n)
    Dim i As Long, j As Long, k As Long
    ' Build augmented [A | I]
    For i = 1 To n
        For j = 1 To n
            aug(i, j) = A(i, j)
        Next j
        aug(i, n + i) = 1
    Next i

    Dim pivot As Double, factor As Double
    For i = 1 To n
        ' Partial pivot
        Dim maxRow As Long: maxRow = i
        Dim maxVal As Double: maxVal = Abs(aug(i, i))
        For k = i + 1 To n
            If Abs(aug(k, i)) > maxVal Then
                maxVal = Abs(aug(k, i))
                maxRow = k
            End If
        Next k
        If maxRow <> i Then
            Dim tmp As Double
            For j = 1 To 2 * n
                tmp = aug(i, j): aug(i, j) = aug(maxRow, j): aug(maxRow, j) = tmp
            Next j
        End If
        pivot = aug(i, i)
        If Abs(pivot) < 1E-15 Then
            MatInv = False
            Exit Function
        End If
        ' Scale pivot row
        For j = 1 To 2 * n
            aug(i, j) = aug(i, j) / pivot
        Next j
        ' Eliminate column
        For k = 1 To n
            If k <> i Then
                factor = aug(k, i)
                For j = 1 To 2 * n
                    aug(k, j) = aug(k, j) - factor * aug(i, j)
                Next j
            End If
        Next k
    Next i
    ' Extract inverse
    ReDim Inv(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            Inv(i, j) = aug(i, n + j)
        Next j
    Next i
    MatInv = True
End Function

' ── Column mean ───────────────────────────────────────────────────────────────
Public Function ColMean(X() As Double, col As Long) As Double
    Dim n As Long, i As Long, s As Double
    n = UBound(X, 1)
    For i = 1 To n: s = s + X(i, col): Next i
    ColMean = s / n
End Function

' ── Column sample std dev ─────────────────────────────────────────────────────
Public Function ColStdDev(X() As Double, col As Long, mu As Double) As Double
    Dim n As Long, i As Long, s As Double
    n = UBound(X, 1)
    For i = 1 To n: s = s + (X(i, col) - mu) ^ 2: Next i
    ColStdDev = Sqr(s / (n - 1))
End Function

' ── Standardise a matrix (z-score each column in place) ─────────────────────
Public Sub Standardize(X() As Double, n As Long, p As Long, _
                        means() As Double, stds() As Double)
    Dim i As Long, j As Long
    For j = 1 To p
        means(j) = ColMean(X, j)
        stds(j) = ColStdDev(X, j, means(j))
        If stds(j) < 1E-15 Then stds(j) = 1   ' avoid divide-by-zero
        For i = 1 To n
            X(i, j) = (X(i, j) - means(j)) / stds(j)
        Next i
    Next j
End Sub

' ── Covariance matrix (p x p) ─────────────────────────────────────────────────
Public Function CovMatrix(X() As Double, n As Long, p As Long, _
                           useCorrForm As Boolean) As Double()
    Dim C() As Double
    ReDim C(1 To p, 1 To p)
    Dim mu() As Double: ReDim mu(1 To p)
    Dim i As Long, j As Long, k As Long, s As Double
    For j = 1 To p
        mu(j) = ColMean(X, j)
    Next j
    For i = 1 To p
        For j = i To p
            s = 0
            For k = 1 To n
                s = s + (X(k, i) - mu(i)) * (X(k, j) - mu(j))
            Next k
            C(i, j) = s / (n - 1)
            C(j, i) = C(i, j)
        Next j
    Next i
    ' Convert to correlation if requested
    If useCorrForm Then
        For i = 1 To p
            For j = 1 To p
                If i <> j Then
                    C(i, j) = C(i, j) / (Sqr(C(i, i)) * Sqr(C(j, j)))
                End If
            Next j
        Next i
        For i = 1 To p
            C(i, i) = 1
        Next i
    End If
    CovMatrix = C
End Function

' ── t-distribution two-tailed p-value (approximation: Abramowitz & Stegun) ───
' Valid for df >= 1.
Public Function TDistPValue(tStat As Double, df As Long) As Double
    Dim x As Double, p As Double
    x = Abs(tStat)
    ' Use WorksheetFunction.TDist if available (Excel 2010+)
    On Error GoTo fallback
    TDistPValue = Application.WorksheetFunction.TDist(x, df, 2)
    Exit Function
fallback:
    ' Cornish–Fisher expansion fallback
    Dim z As Double
    z = x * (1 - 1 / (4 * df)) / Sqr(1 + x * x / (2 * df))
    p = 2 * (1 - NormCDF(z))
    TDistPValue = p
End Function

' ── Standard Normal CDF (Hart 1968 rational approximation) ───────────────────
Public Function NormCDF(z As Double) As Double
    Dim p As Double, b() As Double, t As Double
    If z >= 0 Then
        b = Array(0, 0.2316419, 0.319381530, -0.356563782, _
                  1.781477937, -1.821255978, 1.330274429)
        t = 1 / (1 + 0.2316419 * z)
        p = 1 - (1 / Sqr(2 * Application.WorksheetFunction.Pi())) * Exp(-z * z / 2) * _
            (b(2) * t + b(3) * t ^ 2 + b(4) * t ^ 3 + b(5) * t ^ 4 + b(6) * t ^ 5)
        NormCDF = p
    Else
        NormCDF = 1 - NormCDF(-z)
    End If
End Function

' ── t critical value (inverse CDF approximation via bisection) ───────────────
Public Function TInv(alpha As Double, df As Long) As Double
    ' alpha is two-tailed; e.g. alpha=0.05 -> 95% CI
    Dim lo As Double, hi As Double, mid As Double
    lo = 0: hi = 100
    Dim i As Integer
    For i = 1 To 100
        mid = (lo + hi) / 2
        If TDistPValue(mid, df) > alpha Then
            lo = mid
        Else
            hi = mid
        End If
    Next i
    TInv = mid
End Function

' ── Read DATA sheet into a 2D array (rows=obs, cols=variables) ────────────────
' col_start: first data column (2 = Dep Var, 3 = Factor1 ...)
' Returns True on success; fills dates(), Y(), X()
Public Function ReadDataSheet(dates() As Variant, Y() As Double, X() As Double, _
                               n As Long, p As Long) As Boolean
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Sheets("DATA")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")

    Dim startRow As Long, endRow As Long
    startRow = CLng(wsC.Range("DATA_START_ROW").Value)
    endRow   = CLng(wsC.Range("DATA_END_ROW").Value)
    n = endRow - startRow + 1
    If n < 3 Then
        MsgBox "Not enough data rows (need >= 3). Check CONFIG rows 10-11.", vbCritical
        ReadDataSheet = False: Exit Function
    End If

    ' Count non-empty factor columns (C onwards, row 4)
    p = 0
    Dim col As Long
    For col = 3 To 30
        If wsD.Cells(4, col).Value = "" Then Exit For
        p = p + 1
    Next col
    If p = 0 Then
        MsgBox "No factor columns found in DATA row 4 (columns C+).", vbCritical
        ReadDataSheet = False: Exit Function
    End If

    ReDim dates(1 To n)
    ReDim Y(1 To n)
    ReDim X(1 To n, 1 To p)

    Dim r As Long, i As Long
    For r = startRow To endRow
        i = r - startRow + 1
        dates(i) = wsD.Cells(r, 1).Value
        Y(i)     = CDbl(wsD.Cells(r, 2).Value)
        Dim j As Long
        For j = 1 To p
            X(i, j) = CDbl(wsD.Cells(r, j + 2).Value)
        Next j
    Next r
    ReadDataSheet = True
End Function

' ── Write a 2D array to a worksheet range ─────────────────────────────────────
Public Sub WriteArray2D(ws As Worksheet, startRow As Long, startCol As Long, _
                         arr() As Double)
    Dim r As Long, c As Long
    Dim nR As Long, nC As Long
    nR = UBound(arr, 1): nC = UBound(arr, 2)
    For r = 1 To nR
        For c = 1 To nC
            ws.Cells(startRow + r - 1, startCol + c - 1).Value = arr(r, c)
        Next c
    Next r
End Sub

' ── Bubble-sort eigenvalues descending; mirror permutation in eigenvec matrix ─
Public Sub SortEigenDesc(eigenVals() As Double, eigenVecs() As Double, n As Long)
    Dim i As Long, j As Long
    Dim tmpVal As Double, tmpVec As Double
    For i = 1 To n - 1
        For j = 1 To n - i
            If eigenVals(j) < eigenVals(j + 1) Then
                ' Swap eigenvalues
                tmpVal = eigenVals(j): eigenVals(j) = eigenVals(j + 1): eigenVals(j + 1) = tmpVal
                ' Swap columns in eigenvector matrix
                Dim k As Long
                For k = 1 To n
                    tmpVec = eigenVecs(k, j)
                    eigenVecs(k, j) = eigenVecs(k, j + 1)
                    eigenVecs(k, j + 1) = tmpVec
                Next k
            End If
        Next j
    Next i
End Sub

' ── Sign-flip eigenvectors so largest-magnitude element is positive ───────────
' (Makes loadings deterministic across runs.)
Public Sub FixEigenSigns(eigenVecs() As Double, n As Long)
    Dim j As Long, i As Long
    Dim maxAbs As Double, maxIdx As Long
    For j = 1 To n
        maxAbs = 0
        For i = 1 To n
            If Abs(eigenVecs(i, j)) > maxAbs Then
                maxAbs = Abs(eigenVecs(i, j))
                maxIdx = i
            End If
        Next i
        If eigenVecs(maxIdx, j) < 0 Then
            For i = 1 To n: eigenVecs(i, j) = -eigenVecs(i, j): Next i
        End If
    Next j
End Sub
