Attribute VB_Name = "modStorage"
Option Explicit

' ============================================================
' Storage Module  -  Hidden worksheets as the persistence layer
' ============================================================
' Layout of the _INDEX sheet (hidden):
'   Col A: Session ID  (e.g. "PCA_001", "REG_002")
'   Col B: Type        ("PCA" | "REG")
'   Col C: Name        (user-provided friendly name)
'   Col D: Date        (ISO datetime)
'   Col E: Info        (brief summary string)
'
' Each session gets its own hidden sheet named after its ID.
'
' PCA session sheet layout:
'   R1:  "Name" | <name> | <date> | "" | ""
'   R2:  "Stats"| nObs  | nVars  | nComp | stdz(T/F)
'   R3:  "Means"| m1    | m2     | ... (nVars cols)
'   R4:  "VarNames" | v1 | v2 | ...
'   R5:  "PCNames"  | pc1 | pc2 | ...
'   R6:  "Eigenvals"| e1  | e2  | ...
'   R7:  "VarPct"  | p1  | p2  | ...
'   R8:  "CumPct"  | c1  | c2  | ...
'   R9:  "StdDevs" | s1  | s2  | ... (nVars cols)
'   R10..(9+nObs): "Scores" | sc(i,1) | sc(i,2) | ...
'   R(10+nObs): "Loadings" separator
'   R(11+nObs)..(10+nObs+nVars): varname | l(v,1) | l(v,2) | ...
'
' REG session sheet layout:
'   R1:  "Name" | <name> | <date> | <yName> | <pcaSID>
'   R2:  "Stats"| nObs  | k | intercept(T/F) | ""
'   R3:  "Fit"  | Rsq   | AdjRsq | Fstat | Fpval
'   R4:  "DFreg"| df_reg | df_err | ""
'   R5:  "VarName" | "Coef" | "SE" | "t" | "p-val"   (headers)
'   R6..(5+k): varName | beta | se | t | pval
'   R(6+k)..: "Residuals" separator, then residual values

Private Const INDEX_SHEET As String = "_INDEX"
Private Const SHEET_PREFIX As String = "_S_"

' Ensure the index sheet exists
Sub InitStorage()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(INDEX_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = INDEX_SHEET
        ws.Visible = xlSheetVeryHidden
        ws.Cells(1, 1).Value = "SessionID"
        ws.Cells(1, 2).Value = "Type"
        ws.Cells(1, 3).Value = "Name"
        ws.Cells(1, 4).Value = "Date"
        ws.Cells(1, 5).Value = "Info"
        ws.Rows(1).Font.Bold = True
    End If
End Sub

' Generate next sequential ID for PCA or REG
Function GetNewID(idType As String) As String
    Dim ws As Worksheet: Set ws = GetIndexSheet()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim maxNum As Long: maxNum = 0
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = idType Then
            Dim parts() As String: parts = Split(CStr(ws.Cells(i, 1).Value), "_")
            If UBound(parts) >= 1 Then
                Dim num As Long: num = CLng(parts(1))
                If num > maxNum Then maxNum = num
            End If
        End If
    Next i
    GetNewID = idType & "_" & Format(maxNum + 1, "000")
End Function

' Return the index sheet
Function GetIndexSheet() As Worksheet
    Call InitStorage
    Set GetIndexSheet = ThisWorkbook.Sheets(INDEX_SHEET)
End Function

' Return a session sheet (Nothing if not found)
Function GetSessionSheet(sessionID As String) As Worksheet
    Dim shName As String: shName = SHEET_PREFIX & sessionID
    On Error Resume Next
    Set GetSessionSheet = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
End Function

' Add or update a row in the index
Private Sub UpsertIndex(sessionID As String, idType As String, sName As String, sDate As String, info As String)
    Dim ws As Worksheet: Set ws = GetIndexSheet()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim targetRow As Long: targetRow = 0
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = sessionID Then targetRow = i: Exit For
    Next i
    If targetRow = 0 Then targetRow = lastRow + 1
    ws.Cells(targetRow, 1).Value = sessionID
    ws.Cells(targetRow, 2).Value = idType
    ws.Cells(targetRow, 3).Value = sName
    ws.Cells(targetRow, 4).Value = sDate
    ws.Cells(targetRow, 5).Value = info
End Sub

' Update a specific info field in the index (used by RenamePC etc.)
Sub UpdateIndexEntry(sessionID As String, field As String, value As String)
    Dim ws As Worksheet: Set ws = GetIndexSheet()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = sessionID Then
            If field = "Name" Then ws.Cells(i, 3).Value = value
            If field = "Info" Then ws.Cells(i, 5).Value = value
            If field = "PCNames" Then ws.Cells(i, 5).Value = value  ' reuse info col for now
            Exit For
        End If
    Next i
End Sub

' Persist a PCA session
Sub SavePCAData(sessionID As String, sName As String, nObs As Long, nVars As Long, nComp As Long, _
                stdz As Boolean, varNames As Variant, pcNames As Variant, _
                evalues As Variant, evecs As Variant, scores As Variant, _
                means As Variant, sds As Variant, varPct As Variant, cumPct As Variant)

    Dim shName As String: shName = SHEET_PREFIX & sessionID
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = shName
    Else
        ws.Cells.ClearContents
    End If
    ws.Visible = xlSheetVeryHidden

    Dim sDate As String: sDate = Format(Now(), "yyyy-mm-dd hh:mm")
    Dim j As Long, i As Long

    ws.Cells(1, 1) = "Name":     ws.Cells(1, 2) = sName:  ws.Cells(1, 3) = sDate
    ws.Cells(2, 1) = "Stats":    ws.Cells(2, 2) = nObs:   ws.Cells(2, 3) = nVars
    ws.Cells(2, 4) = nComp:      ws.Cells(2, 5) = IIf(stdz, "TRUE", "FALSE")
    ws.Cells(3, 1) = "Means"
    For j = 1 To nVars: ws.Cells(3, j + 1) = means(j): Next j
    ws.Cells(4, 1) = "VarNames"
    For j = 1 To nVars: ws.Cells(4, j + 1) = varNames(j): Next j
    ws.Cells(5, 1) = "PCNames"
    For j = 1 To nComp: ws.Cells(5, j + 1) = pcNames(j): Next j
    ws.Cells(6, 1) = "Eigenvals"
    For j = 1 To nComp: ws.Cells(6, j + 1) = evalues(j): Next j
    ws.Cells(7, 1) = "VarPct"
    For j = 1 To nComp: ws.Cells(7, j + 1) = varPct(j): Next j
    ws.Cells(8, 1) = "CumPct"
    For j = 1 To nComp: ws.Cells(8, j + 1) = cumPct(j): Next j
    ws.Cells(9, 1) = "StdDevs"
    For j = 1 To nVars: ws.Cells(9, j + 1) = sds(j): Next j

    ' Scores block (rows 10 .. 9+nObs)
    For i = 1 To nObs
        ws.Cells(9 + i, 1) = i
        For j = 1 To nComp
            ws.Cells(9 + i, j + 1) = scores(i, j)
        Next j
    Next i

    ' Loadings block (rows 11+nObs .. 10+nObs+nVars)
    Dim loadRow As Long: loadRow = 10 + nObs
    ws.Cells(loadRow, 1) = "Loadings"
    For i = 1 To nVars
        ws.Cells(loadRow + i, 1) = varNames(i)
        For j = 1 To nComp
            ws.Cells(loadRow + i, j + 1) = evecs(i, j)
        Next j
    Next i

    ' Update index
    Dim info As String
    info = nObs & " obs, " & nVars & " vars, " & nComp & " PCs, " & IIf(stdz, "corr", "cov")
    Call UpsertIndex(sessionID, "PCA", sName, sDate, info)
End Sub

' Persist a regression session
Sub SaveREGData(sessionID As String, sName As String, yName As String, _
                nObs As Long, k As Long, hasIntercept As Boolean, _
                Rsq As Double, AdjRsq As Double, Fstat As Double, Fpval As Double, _
                df_reg As Long, df_err As Long, _
                varNames As Variant, beta As Variant, se As Variant, tStat As Variant, pVal As Variant, _
                resid As Variant, pcaSID As String)

    Dim shName As String: shName = SHEET_PREFIX & sessionID
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = shName
    Else
        ws.Cells.ClearContents
    End If
    ws.Visible = xlSheetVeryHidden

    Dim sDate As String: sDate = Format(Now(), "yyyy-mm-dd hh:mm")
    Dim j As Long, i As Long

    ws.Cells(1, 1) = "Name":    ws.Cells(1, 2) = sName:  ws.Cells(1, 3) = sDate
    ws.Cells(1, 4) = yName:     ws.Cells(1, 5) = pcaSID
    ws.Cells(2, 1) = "Stats":   ws.Cells(2, 2) = nObs:  ws.Cells(2, 3) = k
    ws.Cells(2, 4) = IIf(hasIntercept, "TRUE", "FALSE")
    ws.Cells(3, 1) = "Fit":     ws.Cells(3, 2) = Rsq:   ws.Cells(3, 3) = AdjRsq
    ws.Cells(3, 4) = Fstat:     ws.Cells(3, 5) = Fpval
    ws.Cells(4, 1) = "DF":      ws.Cells(4, 2) = df_reg: ws.Cells(4, 3) = df_err

    ' Coefficient table header
    ws.Cells(5, 1) = "Variable": ws.Cells(5, 2) = "Coef"
    ws.Cells(5, 3) = "Std.Err":  ws.Cells(5, 4) = "t-stat": ws.Cells(5, 5) = "p-value"
    ws.Rows(5).Font.Bold = True

    For j = 1 To k
        ws.Cells(5 + j, 1) = varNames(j)
        ws.Cells(5 + j, 2) = beta(j, 1)
        ws.Cells(5 + j, 3) = se(j)
        ws.Cells(5 + j, 4) = tStat(j)
        ws.Cells(5 + j, 5) = pVal(j)
    Next j

    ' Residuals
    Dim residRow As Long: residRow = 6 + k
    ws.Cells(residRow, 1) = "Residuals"
    For i = 1 To nObs
        ws.Cells(residRow + i, 1) = resid(i)
    Next i

    ' Index
    Dim info As String
    info = "Y=" & yName & ", n=" & nObs & ", R2=" & Format(Rsq, "0.000") & _
           ", F=" & Format(Fstat, "0.00") & " (p=" & Format(Fpval, "0.000") & ")"
    Call UpsertIndex(sessionID, "REG", sName, sDate, info)
End Sub

' Return list of all sessions: array of rows, each row (1 To 5) = ID,Type,Name,Date,Info
Function ListSessions() As Variant
    Dim ws As Worksheet: Set ws = GetIndexSheet()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ListSessions = Array()
        Exit Function
    End If
    Dim n As Long: n = lastRow - 1
    Dim result() As String: ReDim result(1 To n, 1 To 5)
    Dim i As Long
    For i = 1 To n
        result(i, 1) = CStr(ws.Cells(i + 1, 1).Value)
        result(i, 2) = CStr(ws.Cells(i + 1, 2).Value)
        result(i, 3) = CStr(ws.Cells(i + 1, 3).Value)
        result(i, 4) = CStr(ws.Cells(i + 1, 4).Value)
        result(i, 5) = CStr(ws.Cells(i + 1, 5).Value)
    Next i
    ListSessions = result
End Function

' Delete a session (index row + hidden sheet)
Sub DeleteSession(sessionID As String)
    ' Remove from index
    Dim ws As Worksheet: Set ws = GetIndexSheet()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = sessionID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
    ' Delete session sheet
    Dim shName As String: shName = SHEET_PREFIX & sessionID
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(shName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub
