Attribute VB_Name = "modSetup"
Option Explicit

' ============================================================
' Setup / Bootstrap
' Run SetupPCATool() ONCE to install all forms and sheets.
'
' Requires: File > Options > Trust Center > Trust Center Settings
'           > Macro Settings > "Trust access to the VBA project object model"
' ============================================================

Public gCurrentPCAID   As String
Public gCurrentREGID   As String
Public gPreselPCAID    As String

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
           "Launch any time via:  Alt+F8  >  ShowMainForm", _
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
        ws.Range("A1").Font.Italic = True: ws.Range("A1").Font.Color = RGB(130, 130, 130)
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

' ── helpers ────────────────────────────────────────────────────────────────
Private Function AOR(vbp As Object, name As String, cType As Long) As Object
    ' Add-or-Replace a VB component
    On Error Resume Next: vbp.VBComponents.Remove vbp.VBComponents(name): On Error GoTo 0
    Dim c As Object: Set c = vbp.VBComponents.Add(cType): c.name = name: Set AOR = c
End Function

Private Function AC(cont As Object, typeName As String, cName As String, _
                    L As Single, T As Single, W As Single, H As Single) As Object
    Dim c As Object: Set c = cont.Controls.Add(typeName)
    c.name = cName: c.Left = L: c.Top = T: c.Width = W: c.Height = H: Set AC = c
End Function

' ══════════════════════════════════════════════════════════════════════════
' frmMain  –  hub: new analysis + session history
' ══════════════════════════════════════════════════════════════════════════
Private Sub BuildForm_Main(vbp As Object)
    Dim comp As Object: Set comp = AOR(vbp, "frmMain", 3)
    Dim f As Object: Set f = comp.Designer
    f.Width = 516: f.Height = 416: f.Caption = "PCA & Regression Analysis Tool"
    f.StartUpPosition = 1

    Dim b As Object, lbl As Object

    ' Title bar
    Set lbl = AC(f, "Forms.Label.1", "lblTitle", 0, 0, 510, 24)
    lbl.Caption = "   PCA & Regression Analysis Tool"
    lbl.Font.Size = 12: lbl.Font.Bold = True
    lbl.BackColor = RGB(31, 73, 125): lbl.ForeColor = RGB(255, 255, 255)

    ' New analysis frame
    Dim frN As Object: Set frN = AC(f, "Forms.Frame.1", "fraNew", 6, 30, 498, 60)
    frN.Caption = "New Analysis"
    Set b = AC(frN, "Forms.CommandButton.1", "btnNewPCA", 6, 16, 186, 27)
    b.Caption = "Run New PCA Analysis..."
    Set b = AC(frN, "Forms.CommandButton.1", "btnNewReg", 204, 16, 186, 27)
    b.Caption = "Run New Regression..."

    ' PCA sessions frame
    ' lstPCA: col 0 = session ID (hidden, w=0), col 1 = name, col 2 = date, col 3 = info
    Dim frP As Object: Set frP = AC(f, "Forms.Frame.1", "fraPCA", 6, 96, 498, 126)
    frP.Caption = "PCA Sessions"
    Dim lstP As Object: Set lstP = AC(frP, "Forms.ListBox.1", "lstPCA", 6, 14, 408, 100)
    lstP.ColumnCount = 4: lstP.ColumnWidths = "0;126;84;192"
    Set b = AC(frP, "Forms.CommandButton.1", "btnOpenPCA", 420, 14, 72, 27)
    b.Caption = "Open"
    Set b = AC(frP, "Forms.CommandButton.1", "btnDelPCA", 420, 47, 72, 27)
    b.Caption = "Delete"

    ' Regression sessions frame
    ' lstReg: same hidden-ID-in-col-0 layout
    Dim frR As Object: Set frR = AC(f, "Forms.Frame.1", "fraReg", 6, 228, 498, 126)
    frR.Caption = "Regression Sessions"
    Dim lstR As Object: Set lstR = AC(frR, "Forms.ListBox.1", "lstReg", 6, 14, 408, 100)
    lstR.ColumnCount = 4: lstR.ColumnWidths = "0;126;84;192"
    Set b = AC(frR, "Forms.CommandButton.1", "btnOpenReg", 420, 14, 72, 27)
    b.Caption = "Open"
    Set b = AC(frR, "Forms.CommandButton.1", "btnDelReg", 420, 47, 72, 27)
    b.Caption = "Delete"

    ' Footer
    Set b = AC(f, "Forms.CommandButton.1", "btnRefresh", 330, 363, 78, 27)
    b.Caption = "Refresh"
    Set b = AC(f, "Forms.CommandButton.1", "btnClose", 414, 363, 78, 27)
    b.Caption = "Close"

    comp.CodeModule.AddFromString Join(Array( _
        "Option Explicit", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    RefreshLists", _
        "End Sub", _
        "", _
        "Sub RefreshLists()", _
        "    lstPCA.Clear: lstReg.Clear", _
        "    Dim sessions As Variant: sessions = ListSessions()", _
        "    If Not IsArray(sessions) Then Exit Sub", _
        "    Dim n As Long: On Error Resume Next: n = UBound(sessions,1): On Error GoTo 0", _
        "    If n < 1 Then Exit Sub", _
        "    Dim i As Long", _
        "    For i = 1 To n", _
        "        Dim sid As String: sid = sessions(i,1)", _
        "        Dim tp As String:  tp  = sessions(i,2)", _
        "        Dim nm As String:  nm  = sessions(i,3)", _
        "        Dim dt As String:  dt  = sessions(i,4)", _
        "        Dim inf As String: inf = sessions(i,5)", _
        "        If tp = ""PCA"" Then", _
        "            lstPCA.AddItem sid", _
        "            lstPCA.List(lstPCA.ListCount-1, 1) = nm", _
        "            lstPCA.List(lstPCA.ListCount-1, 2) = dt", _
        "            lstPCA.List(lstPCA.ListCount-1, 3) = inf", _
        "        ElseIf tp = ""REG"" Then", _
        "            lstReg.AddItem sid", _
        "            lstReg.List(lstReg.ListCount-1, 1) = nm", _
        "            lstReg.List(lstReg.ListCount-1, 2) = dt", _
        "            lstReg.List(lstReg.ListCount-1, 3) = inf", _
        "        End If", _
        "    Next i", _
        "End Sub", _
        "", _
        "Private Sub btnNewPCA_Click()", _
        "    frmPCASetup.Show", _
        "    RefreshLists", _
        "End Sub", _
        "", _
        "Private Sub btnNewReg_Click()", _
        "    gPreselPCAID = """"", _
        "    frmRegressionSetup.Show", _
        "    RefreshLists", _
        "End Sub", _
        "", _
        "Private Sub btnOpenPCA_Click()", _
        "    If lstPCA.ListIndex < 0 Then MsgBox ""Select a PCA session."", vbExclamation: Exit Sub", _
        "    gCurrentPCAID = lstPCA.List(lstPCA.ListIndex, 0)", _
        "    frmPCAResults.Show", _
        "    RefreshLists", _
        "End Sub", _
        "", _
        "Private Sub btnDelPCA_Click()", _
        "    If lstPCA.ListIndex < 0 Then MsgBox ""Select a PCA session."", vbExclamation: Exit Sub", _
        "    Dim sid As String: sid = lstPCA.List(lstPCA.ListIndex, 0)", _
        "    If MsgBox(""Delete "" & sid & ""?"", vbYesNo + vbQuestion) = vbYes Then", _
        "        DeleteSession sid: RefreshLists", _
        "    End If", _
        "End Sub", _
        "", _
        "Private Sub btnOpenReg_Click()", _
        "    If lstReg.ListIndex < 0 Then MsgBox ""Select a regression session."", vbExclamation: Exit Sub", _
        "    gCurrentREGID = lstReg.List(lstReg.ListIndex, 0)", _
        "    frmRegressionResults.Show", _
        "End Sub", _
        "", _
        "Private Sub btnDelReg_Click()", _
        "    If lstReg.ListIndex < 0 Then MsgBox ""Select a regression session."", vbExclamation: Exit Sub", _
        "    Dim sid As String: sid = lstReg.List(lstReg.ListIndex, 0)", _
        "    If MsgBox(""Delete "" & sid & ""?"", vbYesNo + vbQuestion) = vbYes Then", _
        "        DeleteSession sid: RefreshLists", _
        "    End If", _
        "End Sub", _
        "", _
        "Private Sub btnRefresh_Click(): RefreshLists: End Sub", _
        "Private Sub btnClose_Click(): Unload Me: End Sub" _
    ), vbCrLf)
End Sub

' ══════════════════════════════════════════════════════════════════════════
' frmPCASetup  –  configure & run PCA
' ══════════════════════════════════════════════════════════════════════════
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

    comp.CodeModule.AddFromString Join(Array( _
        "Option Explicit", _
        "Private mRng As Range", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    txtName.Text = ""PCA "" & Format(Now(),""yyyy-mm-dd HH:MM"")", _
        "End Sub", _
        "", _
        "Private Sub btnPick_Click()", _
        "    On Error Resume Next", _
        "    Set mRng = Application.InputBox(""Select data range:"",""Select Range"",txtRange.Text,,,,,8)", _
        "    On Error GoTo 0", _
        "    If Not mRng Is Nothing Then txtRange.Text = mRng.Address(External:=True)", _
        "End Sub", _
        "", _
        "Private Sub btnRun_Click()", _
        "    If mRng Is Nothing Then", _
        "        If Trim(txtRange.Text) = """" Then MsgBox ""Select a data range."", vbExclamation: Exit Sub", _
        "        On Error Resume Next: Set mRng = Range(txtRange.Text): On Error GoTo 0", _
        "        If mRng Is Nothing Then MsgBox ""Invalid range address."", vbCritical: Exit Sub", _
        "    End If", _
        "    Dim maxC As Long", _
        "    If Trim(txtComps.Text) <> """" Then", _
        "        If Not IsNumeric(txtComps.Text) Then MsgBox ""Max components must be numeric."", vbExclamation: Exit Sub", _
        "        maxC = CLng(txtComps.Text)", _
        "    End If", _
        "    Dim sn As String: sn = Trim(txtName.Text)", _
        "    If sn = """" Then sn = ""PCA "" & Format(Now(),""HH:MM"")", _
        "    Dim sid As String", _
        "    sid = RunPCA(mRng, chkHeaders.Value, chkStdz.Value, maxC, sn)", _
        "    If sid = """" Then Exit Sub", _
        "    gCurrentPCAID = sid", _
        "    Unload Me", _
        "    frmPCAResults.Show", _
        "End Sub", _
        "", _
        "Private Sub btnCancel_Click(): Unload Me: End Sub" _
    ), vbCrLf)
End Sub

' ══════════════════════════════════════════════════════════════════════════
' frmPCAResults  –  variance table, loadings, scores, rename PCs
' ══════════════════════════════════════════════════════════════════════════
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

    ' Tab 0 – Variance Explained
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

    ' Tab 1 – Loadings
    Dim pg1 As Object: Set pg1 = mp.Pages(1)
    Set lbl = AC(pg1, "Forms.Label.1", "lblLH", 6, 4, 540, 16)
    lbl.Caption = "Variable loadings on each principal component"
    lbl.Font.Bold = True
    Dim lstL As Object: Set lstL = AC(pg1, "Forms.ListBox.1", "lstLoadings", 6, 22, 540, 212)
    lstL.ColumnCount = 10: lstL.ColumnWidths = "102;60;60;60;60;60;60;60;60;60"
    lstL.Font.Name = "Courier New": lstL.Font.Size = 9
    Set b = AC(pg1, "Forms.CommandButton.1", "btnCpLoad", 6, 242, 132, 24): b.Caption = "Copy to Sheet"

    ' Tab 2 – Scores preview
    Dim pg2 As Object: Set pg2 = mp.Pages(2)
    Set lbl = AC(pg2, "Forms.Label.1", "lblSH", 6, 4, 540, 16)
    lbl.Caption = "PC Scores preview (first 20 observations)"
    lbl.Font.Bold = True
    Dim lstS As Object: Set lstS = AC(pg2, "Forms.ListBox.1", "lstScores", 6, 22, 540, 212)
    lstS.ColumnCount = 10: lstS.ColumnWidths = "42;72;72;72;72;72;72;72;72;72"
    lstS.Font.Name = "Courier New": lstS.Font.Size = 9
    Set b = AC(pg2, "Forms.CommandButton.1", "btnExpSc", 6, 242, 162, 24): b.Caption = "Export All Scores to Sheet"

    ' Action bar
    Set b = AC(f, "Forms.CommandButton.1", "btnRegress", 6, 420, 192, 27)
    b.Caption = "Regress Using These PCs"
    Set b = AC(f, "Forms.CommandButton.1", "btnExport", 204, 420, 138, 27)
    b.Caption = "Export All to Sheet"
    Set b = AC(f, "Forms.CommandButton.1", "btnClose", 492, 420, 78, 27)
    b.Caption = "Close"

    comp.CodeModule.AddFromString Join(Array( _
        "Option Explicit", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    If gCurrentPCAID = """" Then MsgBox ""No PCA session selected."", vbCritical: Unload Me: Exit Sub", _
        "    Reload", _
        "End Sub", _
        "", _
        "Sub Reload()", _
        "    Dim sid As String: sid = gCurrentPCAID", _
        "    Dim meta As Variant: meta = GetPCAMeta(sid)", _
        "    Me.Caption = ""PCA Results  –  "" & meta(1)", _
        "    lblID.Caption = sid & ""  |  "" & meta(3) & "" obs, "" & meta(4) & "" vars, "" & meta(5) & "" PCs  |  "" & IIf(meta(6)=""TRUE"",""Correlation"",""Covariance"") & "" matrix""", _
        "    ' Variance tab", _
        "    lstVariance.Clear", _
        "    Dim vtbl As Variant: vtbl = GetPCAVarianceTable(sid)", _
        "    Dim i As Long", _
        "    For i = 1 To UBound(vtbl,1)", _
        "        lstVariance.AddItem vtbl(i,1)", _
        "        lstVariance.List(lstVariance.ListCount-1, 1) = vtbl(i,2)", _
        "        lstVariance.List(lstVariance.ListCount-1, 2) = vtbl(i,3)", _
        "        lstVariance.List(lstVariance.ListCount-1, 3) = vtbl(i,4)", _
        "    Next i", _
        "    ' Loadings tab", _
        "    lstLoadings.Clear", _
        "    Dim ltbl As Variant: ltbl = GetPCALoadingsTable(sid)", _
        "    Dim nPC As Long: nPC = UBound(ltbl,2) - 1", _
        "    lstLoadings.ColumnCount = nPC + 1", _
        "    For i = 1 To UBound(ltbl,1)", _
        "        lstLoadings.AddItem ltbl(i,1)", _
        "        Dim j As Long", _
        "        For j = 1 To nPC: lstLoadings.List(lstLoadings.ListCount-1, j) = ltbl(i, j+1): Next j", _
        "    Next i", _
        "    ' Scores tab (preview 20 rows)", _
        "    lstScores.Clear", _
        "    Dim sc As Variant: sc = GetPCScores(sid)", _
        "    lstScores.ColumnCount = UBound(sc,2)", _
        "    Dim maxR As Long: maxR = UBound(sc,1): If maxR > 20 Then maxR = 20", _
        "    For i = 1 To maxR", _
        "        lstScores.AddItem Format(sc(i,1),""0.0000"")", _
        "        For j = 2 To UBound(sc,2): lstScores.List(lstScores.ListCount-1, j-1) = Format(sc(i,j),""0.0000""): Next j", _
        "    Next i", _
        "End Sub", _
        "", _
        "Private Sub btnRename_Click()", _
        "    If lstVariance.ListIndex < 0 Then MsgBox ""Select a component first."", vbExclamation: Exit Sub", _
        "    Dim newNm As String: newNm = Trim(txtNewName.Text)", _
        "    If newNm = """" Then MsgBox ""Enter a new name."", vbExclamation: Exit Sub", _
        "    RenamePC gCurrentPCAID, lstVariance.ListIndex + 1, newNm", _
        "    txtNewName.Text = """"", _
        "    Reload", _
        "End Sub", _
        "", _
        "Private Sub btnCpLoad_Click()", _
        "    Dim wsNm As String: wsNm = ""Loadings_"" & gCurrentPCAID", _
        "    Dim ws As Worksheet", _
        "    On Error Resume Next: Set ws = ThisWorkbook.Sheets(wsNm): On Error GoTo 0", _
        "    If ws Is Nothing Then", _
        "        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))", _
        "        ws.Name = wsNm", _
        "    Else: ws.Cells.ClearContents: End If", _
        "    Dim pcn As Variant: pcn = GetPCNames(gCurrentPCAID)", _
        "    ws.Cells(1,1).Value = ""Variable"": Dim j As Long", _
        "    For j = 1 To UBound(pcn): ws.Cells(1,j+1).Value = pcn(j): Next j", _
        "    ws.Rows(1).Font.Bold = True", _
        "    Dim ltbl As Variant: ltbl = GetPCALoadingsTable(gCurrentPCAID)", _
        "    Dim i As Long", _
        "    For i = 1 To UBound(ltbl,1)", _
        "        ws.Cells(i+1,1).Value = ltbl(i,1)", _
        "        For j = 1 To UBound(ltbl,2)-1: ws.Cells(i+1,j+1).Value = CDbl(ltbl(i,j+1)): Next j", _
        "    Next i", _
        "    ws.Columns.AutoFit: ws.Activate", _
        "    MsgBox ""Loadings copied to '"" & wsNm & ""'."", vbInformation", _
        "End Sub", _
        "", _
        "Private Sub btnExpSc_Click(): ExportScoresToSheet gCurrentPCAID: End Sub", _
        "", _
        "Private Sub btnRegress_Click()", _
        "    gPreselPCAID = gCurrentPCAID: Unload Me: frmRegressionSetup.Show", _
        "End Sub", _
        "", _
        "Private Sub btnExport_Click(): ExportScoresToSheet gCurrentPCAID: End Sub", _
        "Private Sub btnClose_Click(): Unload Me: End Sub" _
    ), vbCrLf)
End Sub

' ══════════════════════════════════════════════════════════════════════════
' frmRegressionSetup  –  select Y, select PCs or custom X, run
' cboPCA: col 0 = display name, col 1 = session ID (hidden, width=0)
' ══════════════════════════════════════════════════════════════════════════
Private Sub BuildForm_RegressionSetup(vbp As Object)
    Dim comp As Object: Set comp = AOR(vbp, "frmRegressionSetup", 3)
    Dim f As Object: Set f = comp.Designer
    f.Width = 516: f.Height = 490: f.Caption = "Configure Regression"
    f.StartUpPosition = 1

    Dim lbl As Object, b As Object
    Set lbl = AC(f, "Forms.Label.1", "lblTitle", 0, 0, 510, 22)
    lbl.Caption = "  Regression Setup": lbl.Font.Bold = True: lbl.Font.Size = 11
    lbl.BackColor = RGB(31, 73, 125): lbl.ForeColor = RGB(255, 255, 255)

    ' Y frame
    Dim frY As Object: Set frY = AC(f, "Forms.Frame.1", "fraY", 6, 28, 498, 66)
    frY.Caption = "Dependent Variable (Y)"
    Set lbl = AC(frY, "Forms.Label.1", "lblYL", 6, 18, 66, 18): lbl.Caption = "Y Range:"
    AC frY, "Forms.TextBox.1", "txtY", 78, 16, 276, 18
    Set b = AC(frY, "Forms.CommandButton.1", "btnPickY", 360, 14, 30, 20): b.Caption = "..."
    Set lbl = AC(frY, "Forms.Label.1", "lblYHint", 6, 42, 486, 16)
    lbl.Caption = "Select a column — label from the row above is used as the variable name"
    lbl.ForeColor = RGB(100, 100, 100)

    ' X frame
    Dim frX As Object: Set frX = AC(f, "Forms.Frame.1", "fraX", 6, 100, 498, 252)
    frX.Caption = "Independent Variables (X)"
    Dim optP As Object: Set optP = AC(frX, "Forms.OptionButton.1", "optPCA", 6, 16, 246, 18)
    optP.Caption = "Use PCA Component Scores": optP.Value = True
    Dim optR As Object: Set optR = AC(frX, "Forms.OptionButton.1", "optRange", 258, 16, 234, 18)
    optR.Caption = "Use Custom Range"

    ' PCA sub-section
    Set lbl = AC(frX, "Forms.Label.1", "lblSP", 6, 38, 84, 18): lbl.Caption = "PCA Session:"
    Dim cbo As Object: Set cbo = AC(frX, "Forms.ComboBox.1", "cboPCA", 96, 36, 396, 18)
    cbo.ColumnCount = 2: cbo.ColumnWidths = "390;0"  ' col 1 = hidden session ID
    Set lbl = AC(frX, "Forms.Label.1", "lblAv", 6, 62, 180, 16): lbl.Caption = "Available PCs:"
    Set lbl = AC(frX, "Forms.Label.1", "lblSl", 294, 62, 204, 16): lbl.Caption = "Selected PCs (for regression):"
    Dim lstA As Object: Set lstA = AC(frX, "Forms.ListBox.1", "lstAvail", 6, 80, 180, 162)
    Dim lstS As Object: Set lstS = AC(frX, "Forms.ListBox.1", "lstSel", 294, 80, 198, 162)
    Set b = AC(frX, "Forms.CommandButton.1", "btnAddPC", 192, 98, 96, 27): b.Caption = "Add  >>"
    Set b = AC(frX, "Forms.CommandButton.1", "btnRemPC", 192, 131, 96, 27): b.Caption = "<< Remove"

    ' Range sub-section (initially invisible)
    Set lbl = AC(frX, "Forms.Label.1", "lblXR", 6, 38, 66, 18): lbl.Caption = "X Range:": lbl.Visible = False
    AC frX, "Forms.TextBox.1", "txtX", 78, 36, 276, 18
    Dim txX As Object: Set txX = frX.Controls("txtX"): txX.Visible = False
    Set b = AC(frX, "Forms.CommandButton.1", "btnPickX", 360, 34, 30, 20): b.Caption = "...": b.Visible = False
    Dim chkXH As Object: Set chkXH = AC(frX, "Forms.CheckBox.1", "chkXHdr", 6, 60, 300, 18)
    chkXH.Caption = "First row contains variable names": chkXH.Value = True: chkXH.Visible = False

    ' Options frame
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

    comp.CodeModule.AddFromString Join(Array( _
        "Option Explicit", _
        "Private mYRng As Range, mXRng As Range", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    txtRegName.Text = ""Regression "" & Format(Now(),""yyyy-mm-dd HH:MM"")", _
        "    LoadPCASessions", _
        "    If gPreselPCAID <> """" Then", _
        "        Dim i As Long", _
        "        For i = 0 To cboPCA.ListCount - 1", _
        "            If cboPCA.List(i, 1) = gPreselPCAID Then cboPCA.ListIndex = i: Exit For", _
        "        Next i", _
        "        LoadPCNames", _
        "    End If", _
        "End Sub", _
        "", _
        "Sub LoadPCASessions()", _
        "    cboPCA.Clear", _
        "    Dim sessions As Variant: sessions = ListSessions()", _
        "    If Not IsArray(sessions) Then Exit Sub", _
        "    Dim n As Long: On Error Resume Next: n = UBound(sessions,1): On Error GoTo 0", _
        "    If n < 1 Then Exit Sub", _
        "    Dim i As Long", _
        "    For i = 1 To n", _
        "        If sessions(i,2) = ""PCA"" Then", _
        "            cboPCA.AddItem sessions(i,3) & "" ("" & sessions(i,1) & "")""", _
        "            cboPCA.List(cboPCA.ListCount-1, 1) = sessions(i,1)", _
        "        End If", _
        "    Next i", _
        "End Sub", _
        "", _
        "Sub LoadPCNames()", _
        "    lstAvail.Clear: lstSel.Clear", _
        "    If cboPCA.ListIndex < 0 Then Exit Sub", _
        "    Dim sid As String: sid = cboPCA.List(cboPCA.ListIndex, 1)", _
        "    Dim pcn As Variant: pcn = GetPCNames(sid)", _
        "    Dim i As Long", _
        "    For i = 1 To UBound(pcn)", _
        "        lstAvail.AddItem pcn(i)", _
        "        lstAvail.ItemData(lstAvail.ListCount-1) = i  ' 1-based PC index", _
        "    Next i", _
        "End Sub", _
        "", _
        "Private Sub cboPCA_Change(): LoadPCNames: End Sub", _
        "", _
        "Private Sub optPCA_Click()", _
        "    Dim s As Boolean: s = optPCA.Value", _
        "    lblSP.Visible=s: cboPCA.Visible=s: lblAv.Visible=s: lblSl.Visible=s", _
        "    lstAvail.Visible=s: lstSel.Visible=s: btnAddPC.Visible=s: btnRemPC.Visible=s", _
        "    lblXR.Visible=Not s: txtX.Visible=Not s: btnPickX.Visible=Not s: chkXHdr.Visible=Not s", _
        "End Sub", _
        "Private Sub optRange_Click(): optPCA_Click: End Sub", _
        "", _
        "Private Sub btnPickY_Click()", _
        "    On Error Resume Next", _
        "    Set mYRng = Application.InputBox(""Select Y (dependent variable) — single column:"",""Y Range"",txtY.Text,,,,,8)", _
        "    On Error GoTo 0", _
        "    If Not mYRng Is Nothing Then txtY.Text = mYRng.Address(External:=True)", _
        "End Sub", _
        "", _
        "Private Sub btnPickX_Click()", _
        "    On Error Resume Next", _
        "    Set mXRng = Application.InputBox(""Select X (independent variables) range:"",""X Range"",txtX.Text,,,,,8)", _
        "    On Error GoTo 0", _
        "    If Not mXRng Is Nothing Then txtX.Text = mXRng.Address(External:=True)", _
        "End Sub", _
        "", _
        "Private Sub btnAddPC_Click()", _
        "    If lstAvail.ListIndex < 0 Then Exit Sub", _
        "    Dim idx As Long: idx = lstAvail.ListIndex", _
        "    lstSel.AddItem lstAvail.List(idx): lstSel.ItemData(lstSel.ListCount-1) = lstAvail.ItemData(idx)", _
        "    lstAvail.RemoveItem idx", _
        "End Sub", _
        "", _
        "Private Sub btnRemPC_Click()", _
        "    If lstSel.ListIndex < 0 Then Exit Sub", _
        "    Dim idx As Long: idx = lstSel.ListIndex", _
        "    lstAvail.AddItem lstSel.List(idx): lstAvail.ItemData(lstAvail.ListCount-1) = lstSel.ItemData(idx)", _
        "    lstSel.RemoveItem idx", _
        "End Sub", _
        "", _
        "Private Sub btnRunReg_Click()", _
        "    ' Resolve Y range", _
        "    If mYRng Is Nothing Then", _
        "        If Trim(txtY.Text) = """" Then MsgBox ""Select a Y range."", vbExclamation: Exit Sub", _
        "        On Error Resume Next: Set mYRng = Range(txtY.Text): On Error GoTo 0", _
        "        If mYRng Is Nothing Then MsgBox ""Invalid Y range."", vbCritical: Exit Sub", _
        "    End If", _
        "    Dim sn As String: sn = Trim(txtRegName.Text)", _
        "    If sn = """" Then sn = ""Reg "" & Format(Now(),""HH:MM"")", _
        "    Dim sid As String", _
        "    If optPCA.Value Then", _
        "        If cboPCA.ListIndex < 0 Then MsgBox ""Select a PCA session."", vbExclamation: Exit Sub", _
        "        If lstSel.ListCount = 0 Then MsgBox ""Move at least one PC to the Selected list."", vbExclamation: Exit Sub", _
        "        Dim pcaSID As String: pcaSID = cboPCA.List(cboPCA.ListIndex, 1)", _
        "        Dim selIdx() As Long: ReDim selIdx(0 To lstSel.ListCount - 1)", _
        "        Dim j As Long", _
        "        For j = 0 To lstSel.ListCount - 1: selIdx(j) = lstSel.ItemData(j): Next j", _
        "        sid = RunRegressionFromPCA(pcaSID, selIdx, mYRng, chkInt.Value, sn)", _
        "    Else", _
        "        If mXRng Is Nothing Then", _
        "            If Trim(txtX.Text) = """" Then MsgBox ""Select an X range."", vbExclamation: Exit Sub", _
        "            On Error Resume Next: Set mXRng = Range(txtX.Text): On Error GoTo 0", _
        "            If mXRng Is Nothing Then MsgBox ""Invalid X range."", vbCritical: Exit Sub", _
        "        End If", _
        "        sid = RunRegressionFromRange(mXRng, chkXHdr.Value, mYRng, False, chkInt.Value, sn)", _
        "    End If", _
        "    If sid = """" Then Exit Sub", _
        "    gCurrentREGID = sid: Unload Me: frmRegressionResults.Show", _
        "End Sub", _
        "", _
        "Private Sub btnCancelReg_Click(): Unload Me: End Sub" _
    ), vbCrLf)
End Sub

' ══════════════════════════════════════════════════════════════════════════
' frmRegressionResults  –  model summary + coefficient table
' ══════════════════════════════════════════════════════════════════════════
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

    comp.CodeModule.AddFromString Join(Array( _
        "Option Explicit", _
        "", _
        "Private Sub UserForm_Initialize()", _
        "    If gCurrentREGID = """" Then MsgBox ""No regression session."", vbCritical: Unload Me: Exit Sub", _
        "    Dim sid As String: sid = gCurrentREGID", _
        "    Dim meta As Variant: meta = GetREGMeta(sid)", _
        "    Me.Caption = ""Regression Results  –  "" & meta(1)", _
        "    lblID.Caption = sid & ""  |  Y = "" & meta(3) & ""  |  n="" & meta(4) & "", k="" & meta(5)", _
        "    Dim s As String", _
        "    s = ""R-squared      : "" & Format(CDbl(meta(6)),""0.0000"") & vbCrLf", _
        "    s = s & ""Adj. R-squared : "" & Format(CDbl(meta(7)),""0.0000"") & vbCrLf", _
        "    s = s & ""F-statistic    : "" & Format(CDbl(meta(8)),""0.000"") & ""   p-value: "" & Format(CDbl(meta(9)),""0.0000"") & vbCrLf", _
        "    s = s & ""Significance   : *** p<0.001  ** p<0.01  * p<0.05  . p<0.1""", _
        "    txtSum.Text = s", _
        "    Dim tbl As Variant: tbl = GetREGCoeffTable(sid)", _
        "    lstCoeff.Clear", _
        "    Dim i As Long", _
        "    For i = 1 To UBound(tbl,1)", _
        "        Dim nm As String: nm = tbl(i,1)", _
        "        If Len(nm) < 26 Then nm = nm & Space(26 - Len(nm))", _
        "        lstCoeff.AddItem nm", _
        "        lstCoeff.List(lstCoeff.ListCount-1, 1) = tbl(i,2)", _
        "        lstCoeff.List(lstCoeff.ListCount-1, 2) = tbl(i,3)", _
        "        lstCoeff.List(lstCoeff.ListCount-1, 3) = tbl(i,4)", _
        "        lstCoeff.List(lstCoeff.ListCount-1, 4) = tbl(i,5)", _
        "        lstCoeff.List(lstCoeff.ListCount-1, 5) = tbl(i,6)", _
        "    Next i", _
        "End Sub", _
        "", _
        "Private Sub btnExp_Click(): ExportRegressionToSheet gCurrentREGID: End Sub", _
        "", _
        "Private Sub btnNewReg_Click()", _
        "    gPreselPCAID = """": Unload Me: frmRegressionSetup.Show", _
        "End Sub", _
        "", _
        "Private Sub btnClose_Click(): Unload Me: End Sub" _
    ), vbCrLf)
End Sub
