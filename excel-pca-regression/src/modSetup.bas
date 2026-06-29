Attribute VB_Name = "modSetup"
Option Explicit

' ============================================================
' Setup / Bootstrap
' Run SetupPCATool() ONCE to install all forms and sheets.
'
' Requires: File > Options > Trust Center > Trust Center Settings
'           > Macro Settings > "Trust access to the VBA project object model"
' ============================================================

Public gCurrentPCAID As String
Public gCurrentREGID As String
Public gPreselPCAID  As String

' Module-level code builder (avoids 25-continuation limit in Array())
Private mCode As String

Private Sub CL(s As String)
    mCode = mCode & s & vbCrLf
End Sub

Private Function Done() As String
    Done = mCode
    mCode = ""
End Function

' ─────────────────────────────────────────────────────────────────────────────
Sub SetupPCATool()
    If Not CheckVBATrust() Then Exit Sub
    Application.ScreenUpdating = False
    Call InitStorage
    Call EnsureDataSheet
    Dim vbp As Object: Set vbp = ThisWorkbook.VBProject
    Call BuildForm_Main(vbp)
    Call BuildForm_PCASetup(vbp)
    Call BuildForm_PCAResults(vbp)
    Call BuildForm_RegressionSetup(vbp)
    Call BuildForm_RegressionResults(vbp)
    Application.ScreenUpdating = True
    MsgBox "PCA & Regression Tool installed!" & vbCrLf & vbCrLf & _
           "Launch any time:  Alt+F8  >  ShowMainForm", _
           vbInformation, "Setup Complete"
    Call ShowMainForm
End Sub

Private Sub EnsureDataSheet()
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Data"): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Data"
        ws.Range("A1").Value = "Paste your numeric data here (variables as columns, observations as rows)"
        ws.Range("A1").Font.Italic = True
        ws.Range("A1").Font.Color = RGB(130, 130, 130)
    End If
End Sub

Private Function CheckVBATrust() As Boolean
    On Error Resume Next
    Dim t As Object: Set t = ThisWorkbook.VBProject.VBComponents
    CheckVBATrust = (Err.Number = 0)
    On Error GoTo 0
    If Not CheckVBATrust Then
        MsgBox "Cannot access VBA project." & vbCrLf & vbCrLf & _
               "Enable: File > Options > Trust Center > Trust Center Settings" & vbCrLf & _
               "> Macro Settings > ""Trust access to the VBA project object model""" & vbCrLf & _
               "Then re-run SetupPCATool().", vbCritical, "Setup Error"
    End If
End Function

' ── helpers ──────────────────────────────────────────────────────────────────

Private Function AOR(vbp As Object, cName As String, cType As Long) As Object
    On Error Resume Next: vbp.VBComponents.Remove vbp.VBComponents(cName): On Error GoTo 0
    Dim c As Object: Set c = vbp.VBComponents.Add(cType): c.Name = cName: Set AOR = c
End Function

Private Function AC(cont As Object, typeName As String, cName As String, _
                    L As Single, T As Single, W As Single, H As Single) As Object
    Dim c As Object: Set c = cont.Controls.Add(typeName)
    c.Name = cName: c.Left = L: c.Top = T: c.Width = W: c.Height = H: Set AC = c
End Function

' ══════════════════════════════════════════════════════════════════════════════
' frmMain
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildForm_Main(vbp As Object)
    Dim comp As Object: Set comp = AOR(vbp, "frmMain", 3)
    Dim f As Object: Set f = comp.Designer
    f.Width = 516: f.Height = 416: f.Caption = "PCA & Regression Analysis Tool"
    f.StartUpPosition = 1

    Dim lbl As Object, b As Object
    Set lbl = AC(f, "Forms.Label.1", "lblTitle", 0, 0, 510, 24)
    lbl.Caption = "   PCA & Regression Analysis Tool"
    lbl.Font.Size = 12: lbl.Font.Bold = True
    lbl.BackColor = RGB(31, 73, 125): lbl.ForeColor = RGB(255, 255, 255)

    Dim frN As Object: Set frN = AC(f, "Forms.Frame.1", "fraNew", 6, 30, 498, 60)
    frN.Caption = "New Analysis"
    Set b = AC(frN, "Forms.CommandButton.1", "btnNewPCA", 6, 16, 186, 27)
    b.Caption = "Run New PCA Analysis..."
    Set b = AC(frN, "Forms.CommandButton.1", "btnNewReg", 204, 16, 186, 27)
    b.Caption = "Run New Regression..."

    Dim frP As Object: Set frP = AC(f, "Forms.Frame.1", "fraPCA", 6, 96, 498, 126)
    frP.Caption = "PCA Sessions"
    Dim lstP As Object: Set lstP = AC(frP, "Forms.ListBox.1", "lstPCA", 6, 14, 408, 100)
    lstP.ColumnCount = 4: lstP.ColumnWidths = "0;126;84;192"
    Set b = AC(frP, "Forms.CommandButton.1", "btnOpenPCA", 420, 14, 72, 27): b.Caption = "Open"
    Set b = AC(frP, "Forms.CommandButton.1", "btnDelPCA", 420, 47, 72, 27): b.Caption = "Delete"

    Dim frR As Object: Set frR = AC(f, "Forms.Frame.1", "fraReg", 6, 228, 498, 126)
    frR.Caption = "Regression Sessions"
    Dim lstR As Object: Set lstR = AC(frR, "Forms.ListBox.1", "lstReg", 6, 14, 408, 100)
    lstR.ColumnCount = 4: lstR.ColumnWidths = "0;126;84;192"
    Set b = AC(frR, "Forms.CommandButton.1", "btnOpenReg", 420, 14, 72, 27): b.Caption = "Open"
    Set b = AC(frR, "Forms.CommandButton.1", "btnDelReg", 420, 47, 72, 27): b.Caption = "Delete"

    Set b = AC(f, "Forms.CommandButton.1", "btnRefresh", 330, 363, 78, 27): b.Caption = "Refresh"
    Set b = AC(f, "Forms.CommandButton.1", "btnClose", 414, 363, 78, 27): b.Caption = "Close"

    comp.CodeModule.AddFromString CodeFor_Main()
End Sub

Private Function CodeFor_Main() As String
    mCode = ""
    CL "Option Explicit"
    CL ""
    CL "Private Sub UserForm_Initialize()"
    CL "    RefreshLists"
    CL "End Sub"
    CL ""
    CL "Sub RefreshLists()"
    CL "    lstPCA.Clear: lstReg.Clear"
    CL "    Dim sessions As Variant: sessions = ListSessions()"
    CL "    If Not IsArray(sessions) Then Exit Sub"
    CL "    Dim n As Long"
    CL "    On Error Resume Next: n = UBound(sessions,1): On Error GoTo 0"
    CL "    If n < 1 Then Exit Sub"
    CL "    Dim i As Long, sid As String, tp As String, nm As String, dt As String, inf As String"
    CL "    For i = 1 To n"
    CL "        sid = sessions(i,1): tp = sessions(i,2): nm = sessions(i,3)"
    CL "        dt  = sessions(i,4): inf = sessions(i,5)"
    CL "        If tp = ""PCA"" Then"
    CL "            lstPCA.AddItem sid"
    CL "            lstPCA.List(lstPCA.ListCount-1, 1) = nm"
    CL "            lstPCA.List(lstPCA.ListCount-1, 2) = dt"
    CL "            lstPCA.List(lstPCA.ListCount-1, 3) = inf"
    CL "        ElseIf tp = ""REG"" Then"
    CL "            lstReg.AddItem sid"
    CL "            lstReg.List(lstReg.ListCount-1, 1) = nm"
    CL "            lstReg.List(lstReg.ListCount-1, 2) = dt"
    CL "            lstReg.List(lstReg.ListCount-1, 3) = inf"
    CL "        End If"
    CL "    Next i"
    CL "End Sub"
    CL ""
    CL "Private Sub btnNewPCA_Click()"
    CL "    frmPCASetup.Show"
    CL "    RefreshLists"
    CL "End Sub"
    CL ""
    CL "Private Sub btnNewReg_Click()"
    CL "    gPreselPCAID = """": frmRegressionSetup.Show: RefreshLists"
    CL "End Sub"
    CL ""
    CL "Private Sub btnOpenPCA_Click()"
    CL "    If lstPCA.ListIndex < 0 Then MsgBox ""Select a PCA session."", vbExclamation: Exit Sub"
    CL "    gCurrentPCAID = lstPCA.List(lstPCA.ListIndex, 0)"
    CL "    frmPCAResults.Show: RefreshLists"
    CL "End Sub"
    CL ""
    CL "Private Sub btnDelPCA_Click()"
    CL "    If lstPCA.ListIndex < 0 Then MsgBox ""Select a PCA session."", vbExclamation: Exit Sub"
    CL "    Dim sid As String: sid = lstPCA.List(lstPCA.ListIndex, 0)"
    CL "    If MsgBox(""Delete "" & sid & ""?"", vbYesNo + vbQuestion) = vbYes Then"
    CL "        DeleteSession sid: RefreshLists"
    CL "    End If"
    CL "End Sub"
    CL ""
    CL "Private Sub btnOpenReg_Click()"
    CL "    If lstReg.ListIndex < 0 Then MsgBox ""Select a regression session."", vbExclamation: Exit Sub"
    CL "    gCurrentREGID = lstReg.List(lstReg.ListIndex, 0)"
    CL "    frmRegressionResults.Show"
    CL "End Sub"
    CL ""
    CL "Private Sub btnDelReg_Click()"
    CL "    If lstReg.ListIndex < 0 Then MsgBox ""Select a regression session."", vbExclamation: Exit Sub"
    CL "    Dim sid As String: sid = lstReg.List(lstReg.ListIndex, 0)"
    CL "    If MsgBox(""Delete "" & sid & ""?"", vbYesNo + vbQuestion) = vbYes Then"
    CL "        DeleteSession sid: RefreshLists"
    CL "    End If"
    CL "End Sub"
    CL ""
    CL "Private Sub btnRefresh_Click(): RefreshLists: End Sub"
    CL "Private Sub btnClose_Click(): Unload Me: End Sub"
    CodeFor_Main = Done()
End Function

' ══════════════════════════════════════════════════════════════════════════════
' frmPCASetup
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildForm_PCASetup(vbp As Object)
    Dim comp As Object: Set comp = AOR(vbp, "frmPCASetup", 3)
    Dim f As Object: Set f = comp.Designer
    f.Width = 396: f.Height = 300: f.Caption = "Configure PCA Analysis"
    f.StartUpPosition = 1

    Dim lbl As Object, b As Object
    Set lbl = AC(f, "Forms.Label.1", "lblTitle", 0, 0, 390, 22)
    lbl.Caption = "  PCA Analysis Setup": lbl.Font.Bold = True: lbl.Font.Size = 11
    lbl.BackColor = RGB(31, 73, 125): lbl.ForeColor = RGB(255, 255, 255)

    Dim frD As Object: Set frD = AC(f, "Forms.Frame.1", "fraData", 6, 28, 378, 78)
    frD.Caption = "Data"
    Set lbl = AC(frD, "Forms.Label.1", "lblRng", 6, 18, 72, 18): lbl.Caption = "Data Range:"
    AC frD, "Forms.TextBox.1", "txtRange", 84, 16, 204, 18
    Set b = AC(frD, "Forms.CommandButton.1", "btnPick", 294, 14, 30, 20): b.Caption = "..."
    Dim chkH As Object: Set chkH = AC(frD, "Forms.CheckBox.1", "chkHeaders", 6, 42, 366, 18)
    chkH.Caption = "First row contains variable names (headers)": chkH.Value = True

    Dim frO As Object: Set frO = AC(f, "Forms.Frame.1", "fraOpts", 6, 112, 378, 72)
    frO.Caption = "Options"
    Dim chkS As Object: Set chkS = AC(frO, "Forms.CheckBox.1", "chkStdz", 6, 16, 366, 18)
    chkS.Caption = "Standardize variables (correlation matrix — recommended for mixed units)"
    chkS.Value = True
    Set lbl = AC(frO, "Forms.Label.1", "lblComps", 6, 40, 150, 18): lbl.Caption = "Max components to retain:"
    AC frO, "Forms.TextBox.1", "txtComps", 162, 38, 48, 18
    Set lbl = AC(frO, "Forms.Label.1", "lblCN", 216, 40, 156, 18)
    lbl.Caption = "(blank = keep all)": lbl.ForeColor = RGB(100, 100, 100)

    Dim frNm As Object: Set frNm = AC(f, "Forms.Frame.1", "fraName", 6, 190, 378, 40)
    frNm.Caption = "Session"
    Set lbl = AC(frNm, "Forms.Label.1", "lblNm", 6, 14, 90, 18): lbl.Caption = "Analysis name:"
    AC frNm, "Forms.TextBox.1", "txtName", 102, 12, 270, 18

    Set b = AC(f, "Forms.CommandButton.1", "btnRun", 6, 248, 114, 27)
    b.Caption = "Run PCA": b.Default = True
    Set b = AC(f, "Forms.CommandButton.1", "btnCancel", 126, 248, 78, 27)
    b.Caption = "Cancel": b.Cancel = True

    comp.CodeModule.AddFromString CodeFor_PCASetup()
End Sub

Private Function CodeFor_PCASetup() As String
    mCode = ""
    CL "Option Explicit"
    CL "Private mRng As Range"
    CL ""
    CL "Private Sub UserForm_Initialize()"
    CL "    txtName.Text = ""PCA "" & Format(Now(), ""yyyy-mm-dd HH:MM"")"
    CL "End Sub"
    CL ""
    CL "Private Sub btnPick_Click()"
    CL "    On Error Resume Next"
    CL "    Set mRng = Application.InputBox(""Select data range:"", ""Select Range"", txtRange.Text, , , , , 8)"
    CL "    On Error GoTo 0"
    CL "    If Not mRng Is Nothing Then txtRange.Text = mRng.Address(External:=True)"
    CL "End Sub"
    CL ""
    CL "Private Sub btnRun_Click()"
    CL "    If mRng Is Nothing Then"
    CL "        If Trim(txtRange.Text) = """" Then MsgBox ""Select a data range."", vbExclamation: Exit Sub"
    CL "        On Error Resume Next: Set mRng = Range(txtRange.Text): On Error GoTo 0"
    CL "        If mRng Is Nothing Then MsgBox ""Invalid range address."", vbCritical: Exit Sub"
    CL "    End If"
    CL "    Dim maxC As Long"
    CL "    If Trim(txtComps.Text) <> """" Then"
    CL "        If Not IsNumeric(txtComps.Text) Then MsgBox ""Max components must be numeric."", vbExclamation: Exit Sub"
    CL "        maxC = CLng(txtComps.Text)"
    CL "    End If"
    CL "    Dim sn As String: sn = Trim(txtName.Text)"
    CL "    If sn = """" Then sn = ""PCA "" & Format(Now(), ""HH:MM"")"
    CL "    Dim sid As String"
    CL "    sid = RunPCA(mRng, chkHeaders.Value, chkStdz.Value, maxC, sn)"
    CL "    If sid = """" Then Exit Sub"
    CL "    gCurrentPCAID = sid"
    CL "    Unload Me"
    CL "    frmPCAResults.Show"
    CL "End Sub"
    CL ""
    CL "Private Sub btnCancel_Click(): Unload Me: End Sub"
    CodeFor_PCASetup = Done()
End Function

' ══════════════════════════════════════════════════════════════════════════════
' frmPCAResults
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildForm_PCAResults(vbp As Object)
    Dim comp As Object: Set comp = AOR(vbp, "frmPCAResults", 3)
    Dim f As Object: Set f = comp.Designer
    f.Width = 576: f.Height = 510: f.Caption = "PCA Results"
    f.StartUpPosition = 1

    Dim lbl As Object, b As Object
    Set lbl = AC(f, "Forms.Label.1", "lblTitle", 0, 0, 570, 22)
    lbl.Caption = "  PCA Results": lbl.Font.Bold = True: lbl.Font.Size = 11
    lbl.BackColor = RGB(31, 73, 125): lbl.ForeColor = RGB(255, 255, 255)

    Set lbl = AC(f, "Forms.Label.1", "lblID", 6, 25, 558, 16)
    lbl.Caption = "": lbl.ForeColor = RGB(80, 80, 80)

    Dim mp As Object: Set mp = AC(f, "Forms.MultiPage.1", "mpRes", 6, 44, 558, 364)
    mp.Pages(0).Caption = "Variance Explained"
    mp.Pages(1).Caption = "Loadings"
    mp.Pages(2).Caption = "Scores (Preview)"

    Dim pg0 As Object: Set pg0 = mp.Pages(0)
    Set lbl = AC(pg0, "Forms.Label.1", "lblVH", 6, 4, 540, 16)
    lbl.Caption = "Component               Eigenvalue    Var %      Cumulative %"
    lbl.Font.Bold = True: lbl.Font.Name = "Courier New": lbl.Font.Size = 9
    Dim lstV As Object: Set lstV = AC(pg0, "Forms.ListBox.1", "lstVariance", 6, 22, 540, 188)
    lstV.ColumnCount = 4: lstV.ColumnWidths = "162;90;90;90"
    lstV.Font.Name = "Courier New": lstV.Font.Size = 9
    Set lbl = AC(pg0, "Forms.Label.1", "lblRen", 6, 218, 96, 18): lbl.Caption = "Rename selected:"
    AC pg0, "Forms.TextBox.1", "txtNewName", 108, 216, 192, 20
    Set b = AC(pg0, "Forms.CommandButton.1", "btnRename", 306, 214, 96, 24): b.Caption = "Rename PC"

    Dim pg1 As Object: Set pg1 = mp.Pages(1)
    Set lbl = AC(pg1, "Forms.Label.1", "lblLH", 6, 4, 540, 16)
    lbl.Caption = "Variable loadings on each principal component": lbl.Font.Bold = True
    Dim lstL As Object: Set lstL = AC(pg1, "Forms.ListBox.1", "lstLoadings", 6, 22, 540, 212)
    lstL.ColumnCount = 10: lstL.ColumnWidths = "102;60;60;60;60;60;60;60;60;60"
    lstL.Font.Name = "Courier New": lstL.Font.Size = 9
    Set b = AC(pg1, "Forms.CommandButton.1", "btnCpLoad", 6, 242, 132, 24): b.Caption = "Copy to Sheet"

    Dim pg2 As Object: Set pg2 = mp.Pages(2)
    Set lbl = AC(pg2, "Forms.Label.1", "lblSH", 6, 4, 540, 16)
    lbl.Caption = "PC Scores preview (first 20 observations)": lbl.Font.Bold = True
    Dim lstS As Object: Set lstS = AC(pg2, "Forms.ListBox.1", "lstScores", 6, 22, 540, 212)
    lstS.ColumnCount = 10: lstS.ColumnWidths = "42;72;72;72;72;72;72;72;72;72"
    lstS.Font.Name = "Courier New": lstS.Font.Size = 9
    Set b = AC(pg2, "Forms.CommandButton.1", "btnExpSc", 6, 242, 162, 24): b.Caption = "Export All Scores to Sheet"

    Set b = AC(f, "Forms.CommandButton.1", "btnRegress", 6, 420, 192, 27)
    b.Caption = "Regress Using These PCs"
    Set b = AC(f, "Forms.CommandButton.1", "btnExport", 204, 420, 138, 27): b.Caption = "Export All to Sheet"
    Set b = AC(f, "Forms.CommandButton.1", "btnClose", 492, 420, 78, 27): b.Caption = "Close"

    comp.CodeModule.AddFromString CodeFor_PCAResults()
End Sub

Private Function CodeFor_PCAResults() As String
    mCode = ""
    CL "Option Explicit"
    CL ""
    CL "Private Sub UserForm_Initialize()"
    CL "    If gCurrentPCAID = """" Then MsgBox ""No PCA session selected."", vbCritical: Unload Me: Exit Sub"
    CL "    Reload"
    CL "End Sub"
    CL ""
    CL "Sub Reload()"
    CL "    Dim sid As String: sid = gCurrentPCAID"
    CL "    Dim meta As Variant: meta = GetPCAMeta(sid)"
    CL "    Me.Caption = ""PCA Results  -  "" & meta(1)"
    CL "    lblID.Caption = sid & ""  |  "" & meta(3) & "" obs, "" & meta(4) & "" vars, "" & meta(5) & "" PCs  |  "" & IIf(meta(6) = ""TRUE"", ""Correlation"", ""Covariance"") & "" matrix"""
    CL "    lstVariance.Clear"
    CL "    Dim vtbl As Variant: vtbl = GetPCAVarianceTable(sid)"
    CL "    Dim i As Long"
    CL "    For i = 1 To UBound(vtbl, 1)"
    CL "        lstVariance.AddItem vtbl(i, 1)"
    CL "        lstVariance.List(lstVariance.ListCount - 1, 1) = vtbl(i, 2)"
    CL "        lstVariance.List(lstVariance.ListCount - 1, 2) = vtbl(i, 3)"
    CL "        lstVariance.List(lstVariance.ListCount - 1, 3) = vtbl(i, 4)"
    CL "    Next i"
    CL "    lstLoadings.Clear"
    CL "    Dim ltbl As Variant: ltbl = GetPCALoadingsTable(sid)"
    CL "    Dim nPC As Long: nPC = UBound(ltbl, 2) - 1"
    CL "    lstLoadings.ColumnCount = nPC + 1"
    CL "    Dim j As Long"
    CL "    For i = 1 To UBound(ltbl, 1)"
    CL "        lstLoadings.AddItem ltbl(i, 1)"
    CL "        For j = 1 To nPC"
    CL "            lstLoadings.List(lstLoadings.ListCount - 1, j) = ltbl(i, j + 1)"
    CL "        Next j"
    CL "    Next i"
    CL "    lstScores.Clear"
    CL "    Dim sc As Variant: sc = GetPCScores(sid)"
    CL "    lstScores.ColumnCount = UBound(sc, 2)"
    CL "    Dim maxR As Long: maxR = UBound(sc, 1): If maxR > 20 Then maxR = 20"
    CL "    For i = 1 To maxR"
    CL "        lstScores.AddItem Format(sc(i, 1), ""0.0000"")"
    CL "        For j = 2 To UBound(sc, 2)"
    CL "            lstScores.List(lstScores.ListCount - 1, j - 1) = Format(sc(i, j), ""0.0000"")"
    CL "        Next j"
    CL "    Next i"
    CL "End Sub"
    CL ""
    CL "Private Sub btnRename_Click()"
    CL "    If lstVariance.ListIndex < 0 Then MsgBox ""Select a component first."", vbExclamation: Exit Sub"
    CL "    Dim newNm As String: newNm = Trim(txtNewName.Text)"
    CL "    If newNm = """" Then MsgBox ""Enter a new name."", vbExclamation: Exit Sub"
    CL "    RenamePC gCurrentPCAID, lstVariance.ListIndex + 1, newNm"
    CL "    txtNewName.Text = """": Reload"
    CL "End Sub"
    CL ""
    CL "Private Sub btnCpLoad_Click()"
    CL "    Dim wsNm As String: wsNm = ""Loadings_"" & gCurrentPCAID"
    CL "    Dim ws As Worksheet"
    CL "    On Error Resume Next: Set ws = ThisWorkbook.Sheets(wsNm): On Error GoTo 0"
    CL "    If ws Is Nothing Then"
    CL "        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))"
    CL "        ws.Name = wsNm"
    CL "    Else: ws.Cells.ClearContents: End If"
    CL "    Dim pcn As Variant: pcn = GetPCNames(gCurrentPCAID)"
    CL "    ws.Cells(1, 1).Value = ""Variable"""
    CL "    Dim j As Long"
    CL "    For j = 1 To UBound(pcn): ws.Cells(1, j + 1).Value = pcn(j): Next j"
    CL "    ws.Rows(1).Font.Bold = True"
    CL "    Dim ltbl As Variant: ltbl = GetPCALoadingsTable(gCurrentPCAID)"
    CL "    Dim i As Long"
    CL "    For i = 1 To UBound(ltbl, 1)"
    CL "        ws.Cells(i + 1, 1).Value = ltbl(i, 1)"
    CL "        For j = 1 To UBound(ltbl, 2) - 1: ws.Cells(i + 1, j + 1).Value = CDbl(ltbl(i, j + 1)): Next j"
    CL "    Next i"
    CL "    ws.Columns.AutoFit: ws.Activate"
    CL "    MsgBox ""Loadings copied to '"" & wsNm & ""'."", vbInformation"
    CL "End Sub"
    CL ""
    CL "Private Sub btnExpSc_Click(): ExportScoresToSheet gCurrentPCAID: End Sub"
    CL ""
    CL "Private Sub btnRegress_Click()"
    CL "    gPreselPCAID = gCurrentPCAID: Unload Me: frmRegressionSetup.Show"
    CL "End Sub"
    CL ""
    CL "Private Sub btnExport_Click(): ExportScoresToSheet gCurrentPCAID: End Sub"
    CL "Private Sub btnClose_Click(): Unload Me: End Sub"
    CodeFor_PCAResults = Done()
End Function

' ══════════════════════════════════════════════════════════════════════════════
' frmRegressionSetup
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildForm_RegressionSetup(vbp As Object)
    Dim comp As Object: Set comp = AOR(vbp, "frmRegressionSetup", 3)
    Dim f As Object: Set f = comp.Designer
    f.Width = 516: f.Height = 490: f.Caption = "Configure Regression"
    f.StartUpPosition = 1

    Dim lbl As Object, b As Object
    Set lbl = AC(f, "Forms.Label.1", "lblTitle", 0, 0, 510, 22)
    lbl.Caption = "  Regression Setup": lbl.Font.Bold = True: lbl.Font.Size = 11
    lbl.BackColor = RGB(31, 73, 125): lbl.ForeColor = RGB(255, 255, 255)

    Dim frY As Object: Set frY = AC(f, "Forms.Frame.1", "fraY", 6, 28, 498, 66)
    frY.Caption = "Dependent Variable (Y)"
    Set lbl = AC(frY, "Forms.Label.1", "lblYL", 6, 18, 66, 18): lbl.Caption = "Y Range:"
    AC frY, "Forms.TextBox.1", "txtY", 78, 16, 276, 18
    Set b = AC(frY, "Forms.CommandButton.1", "btnPickY", 360, 14, 30, 20): b.Caption = "..."
    Set lbl = AC(frY, "Forms.Label.1", "lblYHint", 6, 42, 486, 16)
    lbl.Caption = "Select a column — label from row above used as variable name"
    lbl.ForeColor = RGB(100, 100, 100)

    Dim frX As Object: Set frX = AC(f, "Forms.Frame.1", "fraX", 6, 100, 498, 252)
    frX.Caption = "Independent Variables (X)"
    Dim optP As Object: Set optP = AC(frX, "Forms.OptionButton.1", "optPCA", 6, 16, 246, 18)
    optP.Caption = "Use PCA Component Scores": optP.Value = True
    Dim optR As Object: Set optR = AC(frX, "Forms.OptionButton.1", "optRange", 258, 16, 234, 18)
    optR.Caption = "Use Custom Range"

    Set lbl = AC(frX, "Forms.Label.1", "lblSP", 6, 38, 84, 18): lbl.Caption = "PCA Session:"
    Dim cbo As Object: Set cbo = AC(frX, "Forms.ComboBox.1", "cboPCA", 96, 36, 396, 18)
    cbo.ColumnCount = 2: cbo.ColumnWidths = "390;0"
    Set lbl = AC(frX, "Forms.Label.1", "lblAv", 6, 62, 180, 16): lbl.Caption = "Available PCs:"
    Set lbl = AC(frX, "Forms.Label.1", "lblSl", 294, 62, 204, 16): lbl.Caption = "Selected PCs:"
    AC frX, "Forms.ListBox.1", "lstAvail", 6, 80, 180, 162
    AC frX, "Forms.ListBox.1", "lstSel", 294, 80, 198, 162
    Set b = AC(frX, "Forms.CommandButton.1", "btnAddPC", 192, 98, 96, 27): b.Caption = "Add >>"
    Set b = AC(frX, "Forms.CommandButton.1", "btnRemPC", 192, 131, 96, 27): b.Caption = "<< Remove"

    Set lbl = AC(frX, "Forms.Label.1", "lblXR", 6, 38, 66, 18)
    lbl.Caption = "X Range:": lbl.Visible = False
    Dim txX As Object: Set txX = AC(frX, "Forms.TextBox.1", "txtX", 78, 36, 276, 18)
    txX.Visible = False
    Set b = AC(frX, "Forms.CommandButton.1", "btnPickX", 360, 34, 30, 20)
    b.Caption = "...": b.Visible = False
    Dim chkXH As Object: Set chkXH = AC(frX, "Forms.CheckBox.1", "chkXHdr", 6, 60, 300, 18)
    chkXH.Caption = "First row contains variable names": chkXH.Value = True: chkXH.Visible = False

    Dim frOp As Object: Set frOp = AC(f, "Forms.Frame.1", "fraOp", 6, 358, 498, 54)
    frOp.Caption = "Options"
    Dim chkI As Object: Set chkI = AC(frOp, "Forms.CheckBox.1", "chkInt", 6, 14, 252, 18)
    chkI.Caption = "Include intercept (constant)": chkI.Value = True
    Set lbl = AC(frOp, "Forms.Label.1", "lblSN", 264, 14, 90, 18): lbl.Caption = "Session name:"
    AC frOp, "Forms.TextBox.1", "txtRegName", 360, 12, 132, 18

    Set b = AC(f, "Forms.CommandButton.1", "btnRunReg", 6, 424, 126, 27)
    b.Caption = "Run Regression": b.Default = True
    Set b = AC(f, "Forms.CommandButton.1", "btnCancelReg", 138, 424, 78, 27)
    b.Caption = "Cancel": b.Cancel = True

    comp.CodeModule.AddFromString CodeFor_RegressionSetup()
End Sub

Private Function CodeFor_RegressionSetup() As String
    mCode = ""
    CL "Option Explicit"
    CL "Private mYRng As Range, mXRng As Range"
    CL ""
    CL "Private Sub UserForm_Initialize()"
    CL "    txtRegName.Text = ""Regression "" & Format(Now(), ""yyyy-mm-dd HH:MM"")"
    CL "    LoadPCASessions"
    CL "    If gPreselPCAID <> """" Then"
    CL "        Dim i As Long"
    CL "        For i = 0 To cboPCA.ListCount - 1"
    CL "            If cboPCA.List(i, 1) = gPreselPCAID Then cboPCA.ListIndex = i: Exit For"
    CL "        Next i"
    CL "        LoadPCNames"
    CL "    End If"
    CL "End Sub"
    CL ""
    CL "Sub LoadPCASessions()"
    CL "    cboPCA.Clear"
    CL "    Dim sessions As Variant: sessions = ListSessions()"
    CL "    If Not IsArray(sessions) Then Exit Sub"
    CL "    Dim n As Long"
    CL "    On Error Resume Next: n = UBound(sessions, 1): On Error GoTo 0"
    CL "    If n < 1 Then Exit Sub"
    CL "    Dim i As Long"
    CL "    For i = 1 To n"
    CL "        If sessions(i, 2) = ""PCA"" Then"
    CL "            cboPCA.AddItem sessions(i, 3) & "" ("" & sessions(i, 1) & "")"""
    CL "            cboPCA.List(cboPCA.ListCount - 1, 1) = sessions(i, 1)"
    CL "        End If"
    CL "    Next i"
    CL "End Sub"
    CL ""
    CL "Sub LoadPCNames()"
    CL "    lstAvail.Clear: lstSel.Clear"
    CL "    If cboPCA.ListIndex < 0 Then Exit Sub"
    CL "    Dim sid As String: sid = cboPCA.List(cboPCA.ListIndex, 1)"
    CL "    Dim pcn As Variant: pcn = GetPCNames(sid)"
    CL "    Dim i As Long"
    CL "    For i = 1 To UBound(pcn)"
    CL "        lstAvail.AddItem pcn(i)"
    CL "        lstAvail.ItemData(lstAvail.ListCount - 1) = i"
    CL "    Next i"
    CL "End Sub"
    CL ""
    CL "Private Sub cboPCA_Change(): LoadPCNames: End Sub"
    CL ""
    CL "Private Sub optPCA_Click()"
    CL "    Dim s As Boolean: s = optPCA.Value"
    CL "    lblSP.Visible = s: cboPCA.Visible = s: lblAv.Visible = s: lblSl.Visible = s"
    CL "    lstAvail.Visible = s: lstSel.Visible = s: btnAddPC.Visible = s: btnRemPC.Visible = s"
    CL "    lblXR.Visible = Not s: txtX.Visible = Not s: btnPickX.Visible = Not s: chkXHdr.Visible = Not s"
    CL "End Sub"
    CL "Private Sub optRange_Click(): optPCA_Click: End Sub"
    CL ""
    CL "Private Sub btnPickY_Click()"
    CL "    On Error Resume Next"
    CL "    Set mYRng = Application.InputBox(""Select Y (dependent variable) - single column:"", ""Y Range"", txtY.Text, , , , , 8)"
    CL "    On Error GoTo 0"
    CL "    If Not mYRng Is Nothing Then txtY.Text = mYRng.Address(External:=True)"
    CL "End Sub"
    CL ""
    CL "Private Sub btnPickX_Click()"
    CL "    On Error Resume Next"
    CL "    Set mXRng = Application.InputBox(""Select X (independent variables) range:"", ""X Range"", txtX.Text, , , , , 8)"
    CL "    On Error GoTo 0"
    CL "    If Not mXRng Is Nothing Then txtX.Text = mXRng.Address(External:=True)"
    CL "End Sub"
    CL ""
    CL "Private Sub btnAddPC_Click()"
    CL "    If lstAvail.ListIndex < 0 Then Exit Sub"
    CL "    Dim idx As Long: idx = lstAvail.ListIndex"
    CL "    lstSel.AddItem lstAvail.List(idx)"
    CL "    lstSel.ItemData(lstSel.ListCount - 1) = lstAvail.ItemData(idx)"
    CL "    lstAvail.RemoveItem idx"
    CL "End Sub"
    CL ""
    CL "Private Sub btnRemPC_Click()"
    CL "    If lstSel.ListIndex < 0 Then Exit Sub"
    CL "    Dim idx As Long: idx = lstSel.ListIndex"
    CL "    lstAvail.AddItem lstSel.List(idx)"
    CL "    lstAvail.ItemData(lstAvail.ListCount - 1) = lstSel.ItemData(idx)"
    CL "    lstSel.RemoveItem idx"
    CL "End Sub"
    CL ""
    CL "Private Sub btnRunReg_Click()"
    CL "    If mYRng Is Nothing Then"
    CL "        If Trim(txtY.Text) = """" Then MsgBox ""Select a Y range."", vbExclamation: Exit Sub"
    CL "        On Error Resume Next: Set mYRng = Range(txtY.Text): On Error GoTo 0"
    CL "        If mYRng Is Nothing Then MsgBox ""Invalid Y range."", vbCritical: Exit Sub"
    CL "    End If"
    CL "    Dim sn As String: sn = Trim(txtRegName.Text)"
    CL "    If sn = """" Then sn = ""Reg "" & Format(Now(), ""HH:MM"")"
    CL "    Dim sid As String"
    CL "    If optPCA.Value Then"
    CL "        If cboPCA.ListIndex < 0 Then MsgBox ""Select a PCA session."", vbExclamation: Exit Sub"
    CL "        If lstSel.ListCount = 0 Then MsgBox ""Move at least one PC to the Selected list."", vbExclamation: Exit Sub"
    CL "        Dim pcaSID As String: pcaSID = cboPCA.List(cboPCA.ListIndex, 1)"
    CL "        Dim selIdx() As Long: ReDim selIdx(0 To lstSel.ListCount - 1)"
    CL "        Dim j As Long"
    CL "        For j = 0 To lstSel.ListCount - 1: selIdx(j) = lstSel.ItemData(j): Next j"
    CL "        sid = RunRegressionFromPCA(pcaSID, selIdx, mYRng, chkInt.Value, sn)"
    CL "    Else"
    CL "        If mXRng Is Nothing Then"
    CL "            If Trim(txtX.Text) = """" Then MsgBox ""Select an X range."", vbExclamation: Exit Sub"
    CL "            On Error Resume Next: Set mXRng = Range(txtX.Text): On Error GoTo 0"
    CL "            If mXRng Is Nothing Then MsgBox ""Invalid X range."", vbCritical: Exit Sub"
    CL "        End If"
    CL "        sid = RunRegressionFromRange(mXRng, chkXHdr.Value, mYRng, False, chkInt.Value, sn)"
    CL "    End If"
    CL "    If sid = """" Then Exit Sub"
    CL "    gCurrentREGID = sid: Unload Me: frmRegressionResults.Show"
    CL "End Sub"
    CL ""
    CL "Private Sub btnCancelReg_Click(): Unload Me: End Sub"
    CodeFor_RegressionSetup = Done()
End Function

' ══════════════════════════════════════════════════════════════════════════════
' frmRegressionResults
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildForm_RegressionResults(vbp As Object)
    Dim comp As Object: Set comp = AOR(vbp, "frmRegressionResults", 3)
    Dim f As Object: Set f = comp.Designer
    f.Width = 516: f.Height = 450: f.Caption = "Regression Results"
    f.StartUpPosition = 1

    Dim lbl As Object, b As Object
    Set lbl = AC(f, "Forms.Label.1", "lblTitle", 0, 0, 510, 22)
    lbl.Caption = "  Regression Results": lbl.Font.Bold = True: lbl.Font.Size = 11
    lbl.BackColor = RGB(31, 73, 125): lbl.ForeColor = RGB(255, 255, 255)

    Set lbl = AC(f, "Forms.Label.1", "lblID", 6, 25, 498, 16)
    lbl.Caption = "": lbl.ForeColor = RGB(80, 80, 80)

    Dim frS As Object: Set frS = AC(f, "Forms.Frame.1", "fraSum", 6, 44, 498, 96)
    frS.Caption = "Model Summary"
    Dim txS As Object: Set txS = AC(frS, "Forms.TextBox.1", "txtSum", 6, 14, 486, 76)
    txS.MultiLine = True: txS.Locked = True: txS.ScrollBars = 2
    txS.Font.Name = "Courier New": txS.Font.Size = 9

    Dim frC As Object: Set frC = AC(f, "Forms.Frame.1", "fraCoef", 6, 146, 498, 240)
    frC.Caption = "Coefficients"
    Set lbl = AC(frC, "Forms.Label.1", "lblCH", 6, 14, 486, 16)
    lbl.Caption = "Variable                  Coeff      Std Err    t-stat    p-val    Sig"
    lbl.Font.Name = "Courier New": lbl.Font.Size = 9: lbl.Font.Bold = True
    Dim lstC As Object: Set lstC = AC(frC, "Forms.ListBox.1", "lstCoeff", 6, 32, 486, 200)
    lstC.ColumnCount = 6: lstC.ColumnWidths = "144;72;72;60;60;36"
    lstC.Font.Name = "Courier New": lstC.Font.Size = 9

    Set b = AC(f, "Forms.CommandButton.1", "btnExp", 6, 400, 162, 27): b.Caption = "Export Results to Sheet"
    Set b = AC(f, "Forms.CommandButton.1", "btnNewReg", 174, 400, 162, 27): b.Caption = "Run Another Regression"
    Set b = AC(f, "Forms.CommandButton.1", "btnClose", 432, 400, 78, 27): b.Caption = "Close"

    comp.CodeModule.AddFromString CodeFor_RegressionResults()
End Sub

Private Function CodeFor_RegressionResults() As String
    mCode = ""
    CL "Option Explicit"
    CL ""
    CL "Private Sub UserForm_Initialize()"
    CL "    If gCurrentREGID = """" Then MsgBox ""No regression session."", vbCritical: Unload Me: Exit Sub"
    CL "    Dim sid As String: sid = gCurrentREGID"
    CL "    Dim meta As Variant: meta = GetREGMeta(sid)"
    CL "    Me.Caption = ""Regression Results  -  "" & meta(1)"
    CL "    lblID.Caption = sid & ""  |  Y = "" & meta(3) & ""  |  n="" & meta(4) & "", k="" & meta(5)"
    CL "    Dim s As String"
    CL "    s = ""R-squared      : "" & Format(CDbl(meta(6)), ""0.0000"") & vbCrLf"
    CL "    s = s & ""Adj. R-squared : "" & Format(CDbl(meta(7)), ""0.0000"") & vbCrLf"
    CL "    s = s & ""F-statistic    : "" & Format(CDbl(meta(8)), ""0.000"") & ""   p-value: "" & Format(CDbl(meta(9)), ""0.0000"") & vbCrLf"
    CL "    s = s & ""Significance   : *** p<0.001  ** p<0.01  * p<0.05  . p<0.1"""
    CL "    txtSum.Text = s"
    CL "    Dim tbl As Variant: tbl = GetREGCoeffTable(sid)"
    CL "    lstCoeff.Clear"
    CL "    Dim i As Long, nm As String"
    CL "    For i = 1 To UBound(tbl, 1)"
    CL "        nm = tbl(i, 1)"
    CL "        If Len(nm) < 26 Then nm = nm & Space(26 - Len(nm))"
    CL "        lstCoeff.AddItem nm"
    CL "        lstCoeff.List(lstCoeff.ListCount - 1, 1) = tbl(i, 2)"
    CL "        lstCoeff.List(lstCoeff.ListCount - 1, 2) = tbl(i, 3)"
    CL "        lstCoeff.List(lstCoeff.ListCount - 1, 3) = tbl(i, 4)"
    CL "        lstCoeff.List(lstCoeff.ListCount - 1, 4) = tbl(i, 5)"
    CL "        lstCoeff.List(lstCoeff.ListCount - 1, 5) = tbl(i, 6)"
    CL "    Next i"
    CL "End Sub"
    CL ""
    CL "Private Sub btnExp_Click(): ExportRegressionToSheet gCurrentREGID: End Sub"
    CL ""
    CL "Private Sub btnNewReg_Click()"
    CL "    gPreselPCAID = """": Unload Me: frmRegressionSetup.Show"
    CL "End Sub"
    CL ""
    CL "Private Sub btnClose_Click(): Unload Me: End Sub"
    CodeFor_RegressionResults = Done()
End Function
