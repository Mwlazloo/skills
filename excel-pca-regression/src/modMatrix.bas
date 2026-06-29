Attribute VB_Name = "modMatrix"
Option Explicit

' ============================================================
' Matrix Operations  -  Core linear algebra for PCA & OLS
' ============================================================

' Matrix multiply: C(r1,c2) = A(r1,c1) * B(c1,c2)   (1-based in, 1-based out)
Function MatMul(A As Variant, B As Variant) As Variant
    Dim r1 As Long, c1 As Long, c2 As Long
    Dim i As Long, j As Long, k As Long, s As Double
    r1 = UBound(A, 1): c1 = UBound(A, 2): c2 = UBound(B, 2)
    Dim R() As Double
    ReDim R(1 To r1, 1 To c2)
    For i = 1 To r1
        For j = 1 To c2
            s = 0#
            For k = 1 To c1
                s = s + A(i, k) * B(k, j)
            Next k
            R(i, j) = s
        Next j
    Next i
    MatMul = R
End Function

' Matrix transpose
Function MatTranspose(A As Variant) As Variant
    Dim r As Long, c As Long, i As Long, j As Long
    r = UBound(A, 1): c = UBound(A, 2)
    Dim R() As Double
    ReDim R(1 To c, 1 To r)
    For i = 1 To r
        For j = 1 To c
            R(j, i) = A(i, j)
        Next j
    Next i
    MatTranspose = R
End Function

' Matrix inverse via Gauss-Jordan with partial pivoting
Function MatInverse(A As Variant) As Variant
    Dim n As Long, i As Long, j As Long, k As Long
    Dim pivot As Double, factor As Double, tmp As Double
    n = UBound(A, 1)
    Dim aug() As Double
    ReDim aug(1 To n, 1 To 2 * n)
    For i = 1 To n
        For j = 1 To n
            aug(i, j) = A(i, j)
        Next j
        aug(i, n + i) = 1#
    Next i
    For i = 1 To n
        Dim maxV As Double, maxR As Long
        maxV = Abs(aug(i, i)): maxR = i
        For k = i + 1 To n
            If Abs(aug(k, i)) > maxV Then maxV = Abs(aug(k, i)): maxR = k
        Next k
        If maxR <> i Then
            For j = 1 To 2 * n
                tmp = aug(i, j): aug(i, j) = aug(maxR, j): aug(maxR, j) = tmp
            Next j
        End If
        pivot = aug(i, i)
        If Abs(pivot) < 1E-15 Then
            MsgBox "Singular matrix - cannot invert.", vbCritical
            Exit Function
        End If
        For j = 1 To 2 * n
            aug(i, j) = aug(i, j) / pivot
        Next j
        For k = 1 To n
            If k <> i Then
                factor = aug(k, i)
                For j = 1 To 2 * n
                    aug(k, j) = aug(k, j) - factor * aug(i, j)
                Next j
            End If
        Next k
    Next i
    Dim R() As Double
    ReDim R(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            R(i, j) = aug(i, n + j)
        Next j
    Next i
    MatInverse = R
End Function

' Jacobi eigendecomposition for a real symmetric matrix.
' Returns eigenvalues (desc sorted) in eigenVals() and eigenvectors as columns of eigenVecs().
Sub EigenJacobi(A As Variant, eigenVals As Variant, eigenVecs As Variant)
    Dim n As Long, i As Long, j As Long, k As Long, p As Long, q As Long
    n = UBound(A, 1)
    Dim M() As Double, V() As Double
    ReDim M(1 To n, 1 To n)
    ReDim V(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            M(i, j) = A(i, j)
        Next j
        V(i, i) = 1#
    Next i
    Dim maxIter As Long: maxIter = 200 * n * n
    Dim iter As Long
    For iter = 1 To maxIter
        Dim maxOff As Double: maxOff = 0#
        p = 1: q = 2
        For i = 1 To n - 1
            For j = i + 1 To n
                If Abs(M(i, j)) > maxOff Then maxOff = Abs(M(i, j)): p = i: q = j
            Next j
        Next i
        If maxOff < 1E-14 Then Exit For
        Dim theta As Double, c As Double, s As Double
        Dim diff As Double: diff = M(q, q) - M(p, p)
        If Abs(diff) < 1E-15 Then
            theta = WorksheetFunction.Pi() / 4#
        Else
            theta = 0.5 * Atn(2# * M(p, q) / diff)
        End If
        c = Cos(theta): s = Sin(theta)
        Dim Mpp As Double: Mpp = M(p, p)
        Dim Mqq As Double: Mqq = M(q, q)
        Dim Mpq As Double: Mpq = M(p, q)
        M(p, p) = c * c * Mpp + 2# * s * c * Mpq + s * s * Mqq
        M(q, q) = s * s * Mpp - 2# * s * c * Mpq + c * c * Mqq
        M(p, q) = 0#: M(q, p) = 0#
        For k = 1 To n
            If k <> p And k <> q Then
                Dim Mkp As Double: Mkp = M(k, p)
                Dim Mkq As Double: Mkq = M(k, q)
                M(k, p) = c * Mkp + s * Mkq: M(p, k) = M(k, p)
                M(k, q) = -s * Mkp + c * Mkq: M(q, k) = M(k, q)
            End If
        Next k
        For k = 1 To n
            Dim Vkp As Double: Vkp = V(k, p)
            Dim Vkq As Double: Vkq = V(k, q)
            V(k, p) = c * Vkp + s * Vkq
            V(k, q) = -s * Vkp + c * Vkq
        Next k
    Next iter
    Dim evals() As Double
    ReDim evals(1 To n)
    For i = 1 To n: evals(i) = M(i, i): Next i
    ' Sort descending
    For i = 1 To n - 1
        Dim mxI As Long: mxI = i
        For j = i + 1 To n
            If evals(j) > evals(mxI) Then mxI = j
        Next j
        If mxI <> i Then
            Dim tv As Double: tv = evals(i): evals(i) = evals(mxI): evals(mxI) = tv
            For k = 1 To n
                tv = V(k, i): V(k, i) = V(k, mxI): V(k, mxI) = tv
            Next k
        End If
    Next i
    eigenVals = evals: eigenVecs = V
End Sub

Function ColMeans(data As Variant) As Variant
    Dim r As Long, c As Long, i As Long, j As Long, s As Double
    r = UBound(data, 1): c = UBound(data, 2)
    Dim means() As Double: ReDim means(1 To c)
    For j = 1 To c
        s = 0#
        For i = 1 To r: s = s + data(i, j): Next i
        means(j) = s / r
    Next j
    ColMeans = means
End Function

Function ColStdDevs(data As Variant, means As Variant) As Variant
    Dim r As Long, c As Long, i As Long, j As Long, ss As Double
    r = UBound(data, 1): c = UBound(data, 2)
    Dim sds() As Double: ReDim sds(1 To c)
    For j = 1 To c
        ss = 0#
        For i = 1 To r: ss = ss + (data(i, j) - means(j)) ^ 2: Next i
        If r > 1 Then sds(j) = Sqr(ss / (r - 1)) Else sds(j) = 0#
        If sds(j) < 1E-15 Then sds(j) = 1#
    Next j
    ColStdDevs = sds
End Function

Function StandardizeData(data As Variant, means As Variant, sds As Variant, doScale As Boolean) As Variant
    Dim r As Long, c As Long, i As Long, j As Long
    r = UBound(data, 1): c = UBound(data, 2)
    Dim R() As Double: ReDim R(1 To r, 1 To c)
    For i = 1 To r
        For j = 1 To c
            If doScale Then
                R(i, j) = (data(i, j) - means(j)) / sds(j)
            Else
                R(i, j) = data(i, j) - means(j)
            End If
        Next j
    Next i
    StandardizeData = R
End Function

' Convert a 1-based 2D Variant array (from Range.Value) to 1-based Double array
Function RangeToDouble(v As Variant) As Variant
    Dim r As Long, c As Long, i As Long, j As Long
    r = UBound(v, 1): c = UBound(v, 2)
    Dim R() As Double: ReDim R(1 To r, 1 To c)
    For i = 1 To r
        For j = 1 To c
            If IsNumeric(v(i, j)) Then R(i, j) = CDbl(v(i, j)) Else R(i, j) = 0#
        Next j
    Next i
    RangeToDouble = R
End Function
