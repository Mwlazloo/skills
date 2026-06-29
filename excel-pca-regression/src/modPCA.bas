Attribute VB_Name = "modPCA"
Option Explicit

' ============================================================
' PCA Module  -  Run, store, recall, and rename components
' ============================================================

' Run PCA on a worksheet range and persist the session.
' dataRange  : the cell range containing data (with or without a header row)
' hasHeaders : True if row 1 of the range is variable names
' doStdz     : True = use correlation matrix (standardize); False = covariance matrix
' maxComps   : number of components to retain (0 = keep all)
' sName      : friendly name for this analysis
' Returns    : session ID string (e.g. "PCA_003")
Function RunPCA(dataRange As Range, hasHeaders As Boolean, doStdz As Boolean, _
                maxComps As Long, sName As String) As String

    Application.StatusBar = "Running PCA..."

    Dim raw As Variant: raw = dataRange.Value
    Dim nRow As Long: nRow = UBound(raw, 1)
    Dim nCol As Long: nCol = UBound(raw, 2)

    ' Extract variable names
    Dim varNames() As String: ReDim varNames(1 To nCol)
    Dim dataStart As Long
    If hasHeaders Then
        Dim j As Long
        For j = 1 To nCol: varNames(j) = CStr(raw(1, j)): Next j
        dataStart = 2
    Else
        Dim jj As Long
        For jj = 1 To nCol: varNames(jj) = "Var" & jj: Next jj
        dataStart = 1
    End If

    Dim nObs As Long: nObs = nRow - dataStart + 1
    If nObs < 2 Then
        MsgBox "Need at least 2 observations for PCA.", vbCritical
        Exit Function
    End If

    ' Build numeric data array
    Dim data() As Double: ReDim data(1 To nObs, 1 To nCol)
    Dim i As Long, jd As Long
    For i = 1 To nObs
        For jd = 1 To nCol
            If IsNumeric(raw(i + dataStart - 1, jd)) Then
                data(i, jd) = CDbl(raw(i + dataStart - 1, jd))
            End If
        Next jd
    Next i

    ' Compute means and std devs
    Dim means As Variant: means = ColMeans(data)
    Dim sds As Variant:   sds   = ColStdDevs(data, means)

    ' Standardize
    Dim zData As Variant: zData = StandardizeData(data, means, sds, doStdz)

    ' Covariance matrix of z (already centered/scaled)
    Dim nP As Long: nP = nCol
    Dim Cov() As Double: ReDim Cov(1 To nP, 1 To nP)
    Dim ii As Long, jj2 As Long, kk As Long, s As Double
    For ii = 1 To nP
        For jj2 = ii To nP
            s = 0#
            For kk = 1 To nObs: s = s + zData(kk, ii) * zData(kk, jj2): Next kk
            Cov(ii, jj2) = s / (nObs - 1)
            Cov(jj2, ii) = Cov(ii, jj2)
        Next jj2
    Next ii

    ' Eigendecomposition
    Dim evals As Variant, evecs As Variant
    Call EigenJacobi(Cov, evals, evecs)

    ' Determine number of components to keep
    Dim nComp As Long
    If maxComps <= 0 Or maxComps > nP Then nComp = nP Else nComp = maxComps

    ' Compute scores  (nObs x nComp)
    Dim scores() As Double: ReDim scores(1 To nObs, 1 To nComp)
    Dim totalVar As Double: totalVar = 0#
    Dim ev As Long
    For ev = 1 To nP: totalVar = totalVar + evals(ev): Next ev
    If totalVar < 1E-15 Then totalVar = 1#

    Dim ic As Long, io As Long, iv As Long
    For io = 1 To nObs
        For ic = 1 To nComp
            s = 0#
            For iv = 1 To nP: s = s + zData(io, iv) * evecs(iv, ic): Next iv
            scores(io, ic) = s
        Next ic
    Next io

    ' Default PC names
    Dim pcNames() As String: ReDim pcNames(1 To nComp)
    For ic = 1 To nComp: pcNames(ic) = "PC" & ic: Next ic

    ' Build variance explained
    Dim varPct() As Double:  ReDim varPct(1 To nComp)
    Dim cumPct() As Double:  ReDim cumPct(1 To nComp)
    Dim cumSum As Double: cumSum = 0#
    For ic = 1 To nComp
        varPct(ic) = evals(ic) / totalVar
        cumSum = cumSum + varPct(ic)
        cumPct(ic) = cumSum
    Next ic

    ' Get a new session ID and persist
    Dim sid As String: sid = GetNewID("PCA")
    Call SavePCAData(sid, sName, nObs, nCol, nComp, doStdz, _
                     varNames, pcNames, evals, evecs, scores, means, sds, varPct, cumPct)

    Application.StatusBar = False
    RunPCA = sid
End Function

' Rename a PC within a saved PCA session
Sub RenamePC(sessionID As String, pcIndex As Long, newName As String)
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Sub
    ' Row 5 stores PC names starting at column 2
    ws.Cells(5, pcIndex + 1).Value = newName
    ' Update the index sheet display name
    UpdateIndexEntry sessionID, "PCNames", Join(GetPCNames(sessionID), "|")
End Sub

' Return PC scores as a 2D Double array (nObs x nComp)
Function GetPCScores(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim nObs As Long: nObs = CLng(ws.Cells(2, 2).Value)
    Dim nComp As Long: nComp = CLng(ws.Cells(2, 4).Value)
    ' Scores start at row 10, column 2
    Dim scArr() As Double: ReDim scArr(1 To nObs, 1 To nComp)
    Dim i As Long, j As Long
    For i = 1 To nObs
        For j = 1 To nComp
            scArr(i, j) = CDbl(ws.Cells(9 + i, j + 1).Value)
        Next j
    Next i
    GetPCScores = scArr
End Function

' Return PC names as a String array
Function GetPCNames(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim nComp As Long: nComp = CLng(ws.Cells(2, 4).Value)
    Dim pcn() As String: ReDim pcn(1 To nComp)
    Dim j As Long
    For j = 1 To nComp: pcn(j) = CStr(ws.Cells(5, j + 1).Value): Next j
    GetPCNames = pcn
End Function

' Return variable names from a PCA session
Function GetVarNames(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim nVars As Long: nVars = CLng(ws.Cells(2, 3).Value)
    Dim vn() As String: ReDim vn(1 To nVars)
    Dim j As Long
    For j = 1 To nVars: vn(j) = CStr(ws.Cells(4, j + 1).Value): Next j
    GetVarNames = vn
End Function

' Read session metadata for display
Function GetPCAMeta(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim meta(1 To 7) As String
    meta(1) = CStr(ws.Cells(1, 2).Value)  ' name
    meta(2) = CStr(ws.Cells(1, 3).Value)  ' date
    meta(3) = CStr(ws.Cells(2, 2).Value)  ' nObs
    meta(4) = CStr(ws.Cells(2, 3).Value)  ' nVars
    meta(5) = CStr(ws.Cells(2, 4).Value)  ' nComp
    meta(6) = CStr(ws.Cells(2, 5).Value)  ' standardized
    meta(7) = sessionID
    GetPCAMeta = meta
End Function

' Read eigenvalue / variance explained data for display
' Returns 2D array: rows = components, cols = (pcName, eigenVal, varPct, cumPct)
Function GetPCAVarianceTable(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim nComp As Long: nComp = CLng(ws.Cells(2, 4).Value)
    Dim tbl() As String: ReDim tbl(1 To nComp, 1 To 4)
    Dim j As Long
    For j = 1 To nComp
        tbl(j, 1) = CStr(ws.Cells(5, j + 1).Value)               ' PC name
        tbl(j, 2) = Format(ws.Cells(6, j + 1).Value, "0.0000")   ' eigenvalue
        tbl(j, 3) = Format(ws.Cells(7, j + 1).Value * 100, "0.00") & "%"  ' var %
        tbl(j, 4) = Format(ws.Cells(8, j + 1).Value * 100, "0.00") & "%"  ' cum %
    Next j
    GetPCAVarianceTable = tbl
End Function

' Read loadings: rows = variables, cols = PCs
Function GetPCALoadingsTable(sessionID As String) As Variant
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Function
    Dim nVars As Long: nVars = CLng(ws.Cells(2, 3).Value)
    Dim nComp As Long: nComp = CLng(ws.Cells(2, 4).Value)
    Dim nObs As Long: nObs = CLng(ws.Cells(2, 2).Value)
    ' Scores run rows 10..(9+nObs), loadings separator at (10+nObs), data from (11+nObs)
    Dim loadRow As Long: loadRow = 10 + nObs
    Dim tbl() As String: ReDim tbl(1 To nVars, 1 To nComp + 1)
    Dim i As Long, j As Long
    For i = 1 To nVars
        tbl(i, 1) = CStr(ws.Cells(loadRow + i, 1).Value)  ' var name
        For j = 1 To nComp
            tbl(i, j + 1) = Format(ws.Cells(loadRow + i, j + 1).Value, "0.0000")
        Next j
    Next i
    GetPCALoadingsTable = tbl
End Function

' Export PC scores to the Data sheet for use in regression
Sub ExportScoresToSheet(sessionID As String)
    Dim ws As Worksheet: Set ws = GetSessionSheet(sessionID)
    If ws Is Nothing Then Exit Sub

    Dim nObs As Long: nObs = CLng(ws.Cells(2, 2).Value)
    Dim nComp As Long: nComp = CLng(ws.Cells(2, 4).Value)
    Dim pcNames() As String
    Dim tmp As Variant: tmp = GetPCNames(sessionID)
    Dim j As Long
    ReDim pcNames(1 To nComp)
    For j = 1 To nComp: pcNames(j) = tmp(j): Next j

    ' Find or create a Scores sheet
    Dim scoreWS As Worksheet
    On Error Resume Next
    Set scoreWS = ThisWorkbook.Sheets("PC_Scores")
    On Error GoTo 0
    If scoreWS Is Nothing Then
        Set scoreWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        scoreWS.Name = "PC_Scores"
    End If

    scoreWS.Cells.ClearContents
    scoreWS.Cells(1, 1).Value = "Session: " & sessionID

    ' Headers
    For j = 1 To nComp
        scoreWS.Cells(2, j).Value = pcNames(j)
        scoreWS.Cells(2, j).Font.Bold = True
    Next j

    ' Data
    Dim i As Long
    For i = 1 To nObs
        For j = 1 To nComp
            scoreWS.Cells(2 + i, j).Value = ws.Cells(9 + i, j + 1).Value
        Next j
    Next i

    scoreWS.Activate
    MsgBox "PC scores exported to 'PC_Scores' sheet (rows 3:" & (2 + nObs) & ").", vbInformation
End Sub
