Attribute VB_Name = "INSTALL"
'=============================================================================
' INSTALL.bas  —  Self-bootstrapping VBA installer for PCA Analytics
'
' SETUP STEPS (one time only):
'   1. Open PCA_Analytics.xlsx → File → Save As → PCA_Analytics.xlsm
'   2. In Excel: File → Options → Trust Center → Trust Center Settings →
'      Macro Settings → check "Trust access to the VBA project object model"
'   3. Alt+F11 → File → Import File → select this INSTALL.bas
'   4. In VBA editor press F5  (or Excel: Developer → Macros → InstallModules → Run)
'   5. All modules are created. This INSTALL module is deleted automatically.
'   6. Save the workbook (Ctrl+S).  Done.
'
' After install, use:
'   Ctrl+Shift+A  → Run everything
'   Ctrl+Shift+P  → Run PCA
'   Ctrl+Shift+R  → Run OLS Regression
'   Ctrl+Shift+Q  → Run PCA Regression
'=============================================================================
Option Explicit

Public Sub InstallModules()
    Dim VBProj As Object
    Dim VBComp As Object

    On Error GoTo NoVBEAccess
    Set VBProj = ThisWorkbook.VBProject
    On Error GoTo 0

    Application.ScreenUpdating = False

    ' ── Remove any existing versions of our modules ───────────────────────────
    Dim modNames As Variant
    modNames = Array("modUtils", "modPCA", "modRegression", "modCharts", "modMain")
    Dim nm As Variant
    For Each nm In modNames
        On Error Resume Next
        VBProj.VBComponents.Remove VBProj.VBComponents(CStr(nm))
        On Error GoTo 0
    Next nm

    ' ── Create each module ────────────────────────────────────────────────────
    CreateModUtils VBProj
    CreateModPCA VBProj
    CreateModRegression VBProj
    CreateModCharts VBProj
    CreateModMain VBProj

    ' ── Register keyboard shortcuts ───────────────────────────────────────────
    Application.OnKey "^+P", "RunPCA_Entry"
    Application.OnKey "^+R", "RunRegression_Entry"
    Application.OnKey "^+Q", "RunPCAReg_Entry"
    Application.OnKey "^+A", "RunAll"
    Application.OnKey "^+X", "ResetOutputs"

    Application.ScreenUpdating = True

    ' ── Remove self ───────────────────────────────────────────────────────────
    Dim selfComp As Object
    Set selfComp = VBProj.VBComponents("INSTALL")
    VBProj.VBComponents.Remove selfComp   ' deletes this module

    MsgBox "PCA Analytics installed successfully!" & vbCrLf & vbCrLf & _
           "Modules created: modUtils, modPCA, modRegression, modCharts, modMain" & vbCrLf & vbCrLf & _
           "Keyboard shortcuts registered:" & vbCrLf & _
           "  Ctrl+Shift+A  Run All" & vbCrLf & _
           "  Ctrl+Shift+P  Run PCA" & vbCrLf & _
           "  Ctrl+Shift+R  Run OLS Regression" & vbCrLf & _
           "  Ctrl+Shift+Q  Run PCA Regression" & vbCrLf & _
           "  Ctrl+Shift+X  Reset Outputs" & vbCrLf & vbCrLf & _
           "Save the workbook now (Ctrl+S).", _
           vbInformation, "PCA Analytics — Install Complete"
    Exit Sub

NoVBEAccess:
    Application.ScreenUpdating = True
    MsgBox "Cannot access VBA project." & vbCrLf & vbCrLf & _
           "Fix: File → Options → Trust Center → Trust Center Settings →" & vbCrLf & _
           "Macro Settings → check 'Trust access to the VBA project object model'" & vbCrLf & vbCrLf & _
           "Then re-run InstallModules.", vbCritical, "Install Failed"
End Sub

'=============================================================================
' Each CreateMod* sub adds one module with its full source code
'=============================================================================

Private Sub CreateModUtils(VBProj As Object)
    Dim c As Object
    Set c = VBProj.VBComponents.Add(1)  ' 1 = vbext_ct_StdModule
    c.Name = "modUtils"
    Dim s As String

    s = "Attribute VB_Name = ""modUtils""" & vbCrLf
    s = s & "Option Explicit" & vbCrLf & vbCrLf

    s = s & "Public Function MatMul(A() As Double, B() As Double) As Double()" & vbCrLf
    s = s & "    Dim r As Long, m As Long, c As Long" & vbCrLf
    s = s & "    r = UBound(A, 1): m = UBound(A, 2): c = UBound(B, 2)" & vbCrLf
    s = s & "    Dim C() As Double: ReDim C(1 To r, 1 To c)" & vbCrLf
    s = s & "    Dim i As Long, j As Long, k As Long, sum As Double" & vbCrLf
    s = s & "    For i = 1 To r: For j = 1 To c: sum = 0" & vbCrLf
    s = s & "        For k = 1 To m: sum = sum + A(i,k)*B(k,j): Next k" & vbCrLf
    s = s & "        C(i,j) = sum: Next j: Next i" & vbCrLf
    s = s & "    MatMul = C" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Function MatTrans(A() As Double) As Double()" & vbCrLf
    s = s & "    Dim r As Long, c As Long: r = UBound(A,1): c = UBound(A,2)" & vbCrLf
    s = s & "    Dim T() As Double: ReDim T(1 To c, 1 To r)" & vbCrLf
    s = s & "    Dim i As Long, j As Long" & vbCrLf
    s = s & "    For i = 1 To r: For j = 1 To c: T(j,i) = A(i,j): Next j: Next i" & vbCrLf
    s = s & "    MatTrans = T" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Function MatInv(A() As Double, n As Long, Inv() As Double) As Boolean" & vbCrLf
    s = s & "    Dim aug() As Double: ReDim aug(1 To n, 1 To 2*n)" & vbCrLf
    s = s & "    Dim i As Long, j As Long, k As Long" & vbCrLf
    s = s & "    For i = 1 To n: For j = 1 To n: aug(i,j) = A(i,j): Next j: aug(i,n+i)=1: Next i" & vbCrLf
    s = s & "    Dim piv As Double, fac As Double, tmp As Double" & vbCrLf
    s = s & "    For i = 1 To n" & vbCrLf
    s = s & "        Dim mxR As Long: mxR=i: Dim mxV As Double: mxV=Abs(aug(i,i))" & vbCrLf
    s = s & "        For k=i+1 To n: If Abs(aug(k,i))>mxV Then mxV=Abs(aug(k,i)):mxR=k: End If: Next k" & vbCrLf
    s = s & "        If mxR<>i Then" & vbCrLf
    s = s & "            For j=1 To 2*n: tmp=aug(i,j):aug(i,j)=aug(mxR,j):aug(mxR,j)=tmp: Next j" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        piv = aug(i,i)" & vbCrLf
    s = s & "        If Abs(piv)<1E-15 Then MatInv=False: Exit Function: End If" & vbCrLf
    s = s & "        For j=1 To 2*n: aug(i,j)=aug(i,j)/piv: Next j" & vbCrLf
    s = s & "        For k=1 To n: If k<>i Then" & vbCrLf
    s = s & "            fac=aug(k,i)" & vbCrLf
    s = s & "            For j=1 To 2*n: aug(k,j)=aug(k,j)-fac*aug(i,j): Next j" & vbCrLf
    s = s & "        End If: Next k" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    ReDim Inv(1 To n, 1 To n)" & vbCrLf
    s = s & "    For i=1 To n: For j=1 To n: Inv(i,j)=aug(i,n+j): Next j: Next i" & vbCrLf
    s = s & "    MatInv = True" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Function ColMean(X() As Double, col As Long) As Double" & vbCrLf
    s = s & "    Dim n As Long, i As Long, s As Double: n=UBound(X,1)" & vbCrLf
    s = s & "    For i=1 To n: s=s+X(i,col): Next i: ColMean=s/n" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Function ColStdDev(X() As Double, col As Long, mu As Double) As Double" & vbCrLf
    s = s & "    Dim n As Long, i As Long, s As Double: n=UBound(X,1)" & vbCrLf
    s = s & "    For i=1 To n: s=s+(X(i,col)-mu)^2: Next i" & vbCrLf
    s = s & "    ColStdDev=Sqr(s/(n-1))" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Sub Standardize(X() As Double, n As Long, p As Long, means() As Double, stds() As Double)" & vbCrLf
    s = s & "    Dim i As Long, j As Long" & vbCrLf
    s = s & "    For j=1 To p" & vbCrLf
    s = s & "        means(j)=ColMean(X,j): stds(j)=ColStdDev(X,j,means(j))" & vbCrLf
    s = s & "        If stds(j)<1E-15 Then stds(j)=1" & vbCrLf
    s = s & "        For i=1 To n: X(i,j)=(X(i,j)-means(j))/stds(j): Next i" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Function CovMatrix(X() As Double, n As Long, p As Long) As Double()" & vbCrLf
    s = s & "    Dim C() As Double: ReDim C(1 To p, 1 To p)" & vbCrLf
    s = s & "    Dim mu() As Double: ReDim mu(1 To p)" & vbCrLf
    s = s & "    Dim i As Long, j As Long, k As Long, s As Double" & vbCrLf
    s = s & "    For j=1 To p: mu(j)=ColMean(X,j): Next j" & vbCrLf
    s = s & "    For i=1 To p: For j=i To p: s=0" & vbCrLf
    s = s & "        For k=1 To n: s=s+(X(k,i)-mu(i))*(X(k,j)-mu(j)): Next k" & vbCrLf
    s = s & "        C(i,j)=s/(n-1): C(j,i)=C(i,j): Next j: Next i" & vbCrLf
    s = s & "    CovMatrix=C" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Function CorrMatrix(X() As Double, n As Long, p As Long) As Double()" & vbCrLf
    s = s & "    Dim C() As Double: C = CovMatrix(X, n, p)" & vbCrLf
    s = s & "    Dim i As Long, j As Long" & vbCrLf
    s = s & "    For i=1 To p: For j=1 To p" & vbCrLf
    s = s & "        If i<>j Then C(i,j)=C(i,j)/(Sqr(C(i,i))*Sqr(C(j,j)))" & vbCrLf
    s = s & "    Next j: Next i" & vbCrLf
    s = s & "    For i=1 To p: C(i,i)=1: Next i" & vbCrLf
    s = s & "    CorrMatrix=C" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Function TDistPValue(tStat As Double, df As Long) As Double" & vbCrLf
    s = s & "    On Error GoTo fb" & vbCrLf
    s = s & "    TDistPValue = Application.WorksheetFunction.TDist(Abs(tStat),df,2): Exit Function" & vbCrLf
    s = s & "fb: TDistPValue = 1" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Function TInvCrit(alpha As Double, df As Long) As Double" & vbCrLf
    s = s & "    On Error GoTo fb" & vbCrLf
    s = s & "    TInvCrit = Application.WorksheetFunction.TInv(alpha,df): Exit Function" & vbCrLf
    s = s & "fb: TInvCrit = 1.96" & vbCrLf
    s = s & "End Function" & vbCrLf & vbCrLf

    s = s & "Public Sub SortEigenDesc(eVals() As Double, eVecs() As Double, n As Long)" & vbCrLf
    s = s & "    Dim i As Long, j As Long, k As Long, tmpV As Double, tmpE As Double" & vbCrLf
    s = s & "    For i=1 To n-1: For j=1 To n-i" & vbCrLf
    s = s & "        If eVals(j)<eVals(j+1) Then" & vbCrLf
    s = s & "            tmpV=eVals(j):eVals(j)=eVals(j+1):eVals(j+1)=tmpV" & vbCrLf
    s = s & "            For k=1 To n: tmpE=eVecs(k,j):eVecs(k,j)=eVecs(k,j+1):eVecs(k,j+1)=tmpE: Next k" & vbCrLf
    s = s & "        End If: Next j: Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub FixEigenSigns(eVecs() As Double, n As Long)" & vbCrLf
    s = s & "    Dim i As Long, j As Long, mxA As Double, mxI As Long" & vbCrLf
    s = s & "    For j=1 To n: mxA=0" & vbCrLf
    s = s & "        For i=1 To n: If Abs(eVecs(i,j))>mxA Then mxA=Abs(eVecs(i,j)):mxI=i: End If: Next i" & vbCrLf
    s = s & "        If eVecs(mxI,j)<0 Then For i=1 To n: eVecs(i,j)=-eVecs(i,j): Next i: End If" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub WriteArray2D(ws As Worksheet, sR As Long, sC As Long, arr() As Double, fmt As String)" & vbCrLf
    s = s & "    Dim r As Long, c As Long" & vbCrLf
    s = s & "    For r=1 To UBound(arr,1): For c=1 To UBound(arr,2)" & vbCrLf
    s = s & "        ws.Cells(sR+r-1,sC+c-1).Value=arr(r,c)" & vbCrLf
    s = s & "        If Len(fmt)>0 Then ws.Cells(sR+r-1,sC+c-1).NumberFormat=fmt" & vbCrLf
    s = s & "    Next c: Next r" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Function ReadDataSheet(dates() As Variant, Y() As Double, X() As Double, n As Long, p As Long) As Boolean" & vbCrLf
    s = s & "    Dim wsD As Worksheet: Set wsD=ThisWorkbook.Sheets(""DATA"")" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim sR As Long: sR=CLng(wsC.Range(""DATA_START_ROW"").Value)" & vbCrLf
    s = s & "    Dim eR As Long: eR=CLng(wsC.Range(""DATA_END_ROW"").Value)" & vbCrLf
    s = s & "    n=eR-sR+1" & vbCrLf
    s = s & "    If n<3 Then MsgBox ""Need >=3 data rows. Check CONFIG rows 10-11."",vbCritical: ReadDataSheet=False: Exit Function: End If" & vbCrLf
    s = s & "    p=0: Dim col As Long" & vbCrLf
    s = s & "    For col=3 To 30: If wsD.Cells(4,col).Value="""" Then Exit For: p=p+1: Next col" & vbCrLf
    s = s & "    If p=0 Then MsgBox ""No factor columns in DATA row 4 (cols C+)."",vbCritical: ReadDataSheet=False: Exit Function: End If" & vbCrLf
    s = s & "    ReDim dates(1 To n): ReDim Y(1 To n): ReDim X(1 To n, 1 To p)" & vbCrLf
    s = s & "    Dim r As Long, i As Long, j As Long" & vbCrLf
    s = s & "    For r=sR To eR: i=r-sR+1" & vbCrLf
    s = s & "        dates(i)=wsD.Cells(r,1).Value: Y(i)=CDbl(wsD.Cells(r,2).Value)" & vbCrLf
    s = s & "        For j=1 To p: X(i,j)=CDbl(wsD.Cells(r,j+2).Value): Next j" & vbCrLf
    s = s & "    Next r: ReadDataSheet=True" & vbCrLf
    s = s & "End Function" & vbCrLf

    c.CodeModule.AddFromString s
End Sub

Private Sub CreateModPCA(VBProj As Object)
    Dim c As Object
    Set c = VBProj.VBComponents.Add(1)
    c.Name = "modPCA"
    Dim s As String

    s = "Attribute VB_Name = ""modPCA""" & vbCrLf
    s = s & "Option Explicit" & vbCrLf & vbCrLf
    s = s & "Public g_eigenVals() As Double" & vbCrLf
    s = s & "Public g_eigenVecs() As Double" & vbCrLf
    s = s & "Public g_loadings()  As Double" & vbCrLf
    s = s & "Public g_scores()    As Double" & vbCrLf
    s = s & "Public g_covMat()    As Double" & vbCrLf
    s = s & "Public g_means()     As Double" & vbCrLf
    s = s & "Public g_stds()      As Double" & vbCrLf
    s = s & "Public g_n           As Long" & vbCrLf
    s = s & "Public g_p           As Long" & vbCrLf
    s = s & "Public g_dates()     As Variant" & vbCrLf
    s = s & "Public g_Y()         As Double" & vbCrLf & vbCrLf

    s = s & "Public Sub RunPCA()" & vbCrLf
    s = s & "    Application.ScreenUpdating=False: Application.Calculation=xlCalculationManual" & vbCrLf
    s = s & "    Dim wsP As Worksheet: Set wsP=ThisWorkbook.Sheets(""PCA"")" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim X() As Double" & vbCrLf
    s = s & "    If Not ReadDataSheet(g_dates,g_Y,X,g_n,g_p) Then GoTo Cleanup" & vbCrLf
    s = s & "    Dim useCorrForm As Boolean" & vbCrLf
    s = s & "    useCorrForm=(UCase(Trim(wsC.Range(""MATRIX_TYPE"").Value))=""CORRELATION"")" & vbCrLf
    s = s & "    ReDim g_means(1 To g_p): ReDim g_stds(1 To g_p)" & vbCrLf
    s = s & "    Dim mins() As Double: ReDim mins(1 To g_p)" & vbCrLf
    s = s & "    Dim maxs() As Double: ReDim maxs(1 To g_p)" & vbCrLf
    s = s & "    Dim i As Long, j As Long, r As Long" & vbCrLf
    s = s & "    For j=1 To g_p" & vbCrLf
    s = s & "        g_means(j)=ColMean(X,j): g_stds(j)=ColStdDev(X,j,g_means(j))" & vbCrLf
    s = s & "        Dim mn As Double: mn=X(1,j): Dim mx As Double: mx=X(1,j)" & vbCrLf
    s = s & "        For i=2 To g_n: If X(i,j)<mn Then mn=X(i,j): If X(i,j)>mx Then mx=X(i,j): Next i" & vbCrLf
    s = s & "        mins(j)=mn: maxs(j)=mx" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "    For j=1 To g_p: r=5+j" & vbCrLf
    s = s & "        wsP.Cells(r,2)=g_n: wsP.Cells(r,3)=g_means(j): wsP.Cells(r,4)=g_stds(j)" & vbCrLf
    s = s & "        wsP.Cells(r,5)=mins(j): wsP.Cells(r,6)=maxs(j)" & vbCrLf
    s = s & "        wsP.Cells(r,3).NumberFormat=""0.0000"": wsP.Cells(r,4).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        wsP.Cells(r,5).NumberFormat=""0.0000"": wsP.Cells(r,6).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "    Dim Xstd() As Double: ReDim Xstd(1 To g_n, 1 To g_p)" & vbCrLf
    s = s & "    For i=1 To g_n: For j=1 To g_p: Xstd(i,j)=X(i,j): Next j: Next i" & vbCrLf
    s = s & "    Dim dm1() As Double: ReDim dm1(1 To g_p): Dim dm2() As Double: ReDim dm2(1 To g_p)" & vbCrLf
    s = s & "    If useCorrForm Then Standardize Xstd,g_n,g_p,dm1,dm2" & vbCrLf
    s = s & "    If useCorrForm Then g_covMat=CorrMatrix(Xstd,g_n,g_p) Else g_covMat=CovMatrix(Xstd,g_n,g_p)" & vbCrLf
    s = s & "    Dim matLbl As String: matLbl=IIf(useCorrForm,""CORRELATION"",""COVARIANCE"")" & vbCrLf
    s = s & "    wsP.Cells(17,1).Value=""2.  "" & matLbl & "" MATRIX""" & vbCrLf
    s = s & "    For i=1 To g_p: For j=1 To g_p" & vbCrLf
    s = s & "        wsP.Cells(18+i,j+1)=g_covMat(i,j): wsP.Cells(18+i,j+1).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    Next j: Next i" & vbCrLf
    s = s & "    ReDim g_eigenVals(1 To g_p): ReDim g_eigenVecs(1 To g_p, 1 To g_p)" & vbCrLf
    s = s & "    JacobiEigen g_covMat,g_p,g_eigenVals,g_eigenVecs" & vbCrLf
    s = s & "    SortEigenDesc g_eigenVals,g_eigenVecs,g_p" & vbCrLf
    s = s & "    FixEigenSigns g_eigenVecs,g_p" & vbCrLf
    s = s & "    ReDim g_loadings(1 To g_p, 1 To g_p)" & vbCrLf
    s = s & "    For i=1 To g_p: For j=1 To g_p" & vbCrLf
    s = s & "        g_loadings(i,j)=g_eigenVecs(i,j)*Sqr(IIf(g_eigenVals(j)>0,g_eigenVals(j),0))" & vbCrLf
    s = s & "    Next j: Next i" & vbCrLf
    s = s & "    g_scores=MatMul(Xstd,g_eigenVecs)" & vbCrLf
    s = s & "    Dim totalVar As Double: For j=1 To g_p: totalVar=totalVar+g_eigenVals(j): Next j" & vbCrLf
    s = s & "    Dim cumPct As Double: Dim numPCs As Long: numPCs=CLng(wsC.Range(""NUM_PCS"").Value)" & vbCrLf
    s = s & "    For j=1 To g_p: r=31+j" & vbCrLf
    s = s & "        Dim pctV As Double: pctV=IIf(totalVar>0,g_eigenVals(j)/totalVar,0)" & vbCrLf
    s = s & "        cumPct=cumPct+pctV" & vbCrLf
    s = s & "        wsP.Cells(r,2)=g_eigenVals(j): wsP.Cells(r,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        wsP.Cells(r,3)=pctV:           wsP.Cells(r,3).NumberFormat=""0.00%""" & vbCrLf
    s = s & "        wsP.Cells(r,4)=cumPct:         wsP.Cells(r,4).NumberFormat=""0.00%""" & vbCrLf
    s = s & "        wsP.Cells(r,5)=IIf(j<=numPCs,""Retained"",""Dropped"")" & vbCrLf
    s = s & "        wsP.Cells(r,5).Font.Bold=True" & vbCrLf
    s = s & "        wsP.Cells(r,5).Font.Color=IIf(j<=numPCs,RGB(0,128,0),RGB(180,0,0))" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "    For i=1 To g_p: For j=1 To g_p" & vbCrLf
    s = s & "        wsP.Cells(44+i,j+1)=g_loadings(i,j): wsP.Cells(44+i,j+1).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        wsP.Cells(44+i,j+1).Font.Bold=(Abs(g_loadings(i,j))>=0.5)" & vbCrLf
    s = s & "    Next j: Next i" & vbCrLf
    s = s & "    For i=1 To g_n" & vbCrLf
    s = s & "        wsP.Cells(57+i,1)=g_dates(i): wsP.Cells(57+i,1).NumberFormat=""YYYY-MM-DD""" & vbCrLf
    s = s & "        For j=1 To numPCs: wsP.Cells(57+i,j+1)=g_scores(i,j): wsP.Cells(57+i,j+1).NumberFormat=""0.0000"": Next j" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    BuildPCACharts" & vbCrLf
    s = s & "    Application.Calculation=xlCalculationAutomatic: Application.ScreenUpdating=True" & vbCrLf
    s = s & "    MsgBox ""PCA complete. "" & g_p & "" factors, "" & g_n & "" obs."",vbInformation,""PCA Analytics""" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "Cleanup: Application.Calculation=xlCalculationAutomatic: Application.ScreenUpdating=True" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub JacobiEigen(A() As Double, n As Long, eVals() As Double, eVecs() As Double)" & vbCrLf
    s = s & "    Dim i As Long, j As Long, k As Long" & vbCrLf
    s = s & "    Dim maxIter As Long: maxIter=IIf(200*n*n>500,200*n*n,500)" & vbCrLf
    s = s & "    ReDim eVecs(1 To n, 1 To n)" & vbCrLf
    s = s & "    For i=1 To n: eVecs(i,i)=1: Next i" & vbCrLf
    s = s & "    Dim B() As Double: ReDim B(1 To n, 1 To n)" & vbCrLf
    s = s & "    For i=1 To n: For j=1 To n: B(i,j)=A(i,j): Next j: Next i" & vbCrLf
    s = s & "    Dim p As Long, q As Long, mxOff As Double" & vbCrLf
    s = s & "    Dim theta As Double, c As Double, s As Double" & vbCrLf
    s = s & "    Const TOL As Double = 1E-12" & vbCrLf
    s = s & "    Dim iter As Long" & vbCrLf
    s = s & "    For iter=1 To maxIter" & vbCrLf
    s = s & "        mxOff=0: p=1: q=2" & vbCrLf
    s = s & "        For i=1 To n-1: For j=i+1 To n" & vbCrLf
    s = s & "            If Abs(B(i,j))>mxOff Then mxOff=Abs(B(i,j)):p=i:q=j" & vbCrLf
    s = s & "        Next j: Next i" & vbCrLf
    s = s & "        If mxOff<TOL Then Exit For" & vbCrLf
    s = s & "        Dim denom As Double: denom=B(q,q)-B(p,p)" & vbCrLf
    s = s & "        If Abs(denom)<TOL Then theta=3.14159265358979/4 Else theta=0.5*Atn(2*B(p,q)/denom)" & vbCrLf
    s = s & "        c=Cos(theta): s=Sin(theta)" & vbCrLf
    s = s & "        Dim bpp As Double: bpp=c*c*B(p,p)+2*s*c*B(p,q)+s*s*B(q,q)" & vbCrLf
    s = s & "        Dim bqq As Double: bqq=s*s*B(p,p)-2*s*c*B(p,q)+c*c*B(q,q)" & vbCrLf
    s = s & "        For k=1 To n: If k<>p And k<>q Then" & vbCrLf
    s = s & "            Dim bkp As Double: bkp=c*B(k,p)+s*B(k,q)" & vbCrLf
    s = s & "            Dim bkq As Double: bkq=-s*B(k,p)+c*B(k,q)" & vbCrLf
    s = s & "            B(k,p)=bkp:B(p,k)=bkp:B(k,q)=bkq:B(q,k)=bkq" & vbCrLf
    s = s & "        End If: Next k" & vbCrLf
    s = s & "        B(p,p)=bpp:B(q,q)=bqq:B(p,q)=0:B(q,p)=0" & vbCrLf
    s = s & "        For i=1 To n" & vbCrLf
    s = s & "            Dim bip As Double: bip=eVecs(i,p)" & vbCrLf
    s = s & "            Dim biq As Double: biq=eVecs(i,q)" & vbCrLf
    s = s & "            eVecs(i,p)=c*bip+s*biq: eVecs(i,q)=-s*bip+c*biq" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    Next iter" & vbCrLf
    s = s & "    ReDim eVals(1 To n)" & vbCrLf
    s = s & "    For i=1 To n: eVals(i)=IIf(B(i,i)>0,B(i,i),0): Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf

    c.CodeModule.AddFromString s
End Sub

Private Sub CreateModRegression(VBProj As Object)
    Dim c As Object
    Set c = VBProj.VBComponents.Add(1)
    c.Name = "modRegression"
    Dim s As String

    s = "Attribute VB_Name = ""modRegression""" & vbCrLf
    s = s & "Option Explicit" & vbCrLf & vbCrLf

    s = s & "Public Sub RunRegression()" & vbCrLf
    s = s & "    Application.ScreenUpdating=False: Application.Calculation=xlCalculationManual" & vbCrLf
    s = s & "    Dim Xraw() As Double: Dim dDates() As Variant: Dim dY() As Double" & vbCrLf
    s = s & "    Dim n2 As Long: Dim p2 As Long" & vbCrLf
    s = s & "    If Not ReadDataSheet(dDates,dY,Xraw,n2,p2) Then GoTo Cleanup" & vbCrLf
    s = s & "    OLSCore Xraw,dY,dDates,n2,p2,""REGRESSION"",False" & vbCrLf
    s = s & "    Application.Calculation=xlCalculationAutomatic: Application.ScreenUpdating=True" & vbCrLf
    s = s & "    MsgBox ""OLS Regression complete."",vbInformation,""PCA Analytics"": Exit Sub" & vbCrLf
    s = s & "Cleanup: Application.Calculation=xlCalculationAutomatic: Application.ScreenUpdating=True" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub RunPCARegression()" & vbCrLf
    s = s & "    Application.ScreenUpdating=False: Application.Calculation=xlCalculationManual" & vbCrLf
    s = s & "    If g_n=0 Then RunPCA: If g_n=0 Then GoTo Cleanup" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim wsR As Worksheet: Set wsR=ThisWorkbook.Sheets(""PCA_REG"")" & vbCrLf
    s = s & "    Dim numPCs As Long: numPCs=CLng(wsC.Range(""NUM_PCS"").Value)" & vbCrLf
    s = s & "    Dim inclPC() As Boolean: ReDim inclPC(1 To g_p)" & vbCrLf
    s = s & "    Dim pcCount As Long: Dim j As Long" & vbCrLf
    s = s & "    For j=1 To g_p" & vbCrLf
    s = s & "        inclPC(j)=(UCase(Trim(wsC.Cells(27+j,5).Value))=""YES"")" & vbCrLf
    s = s & "        If inclPC(j) Then pcCount=pcCount+1" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "    If pcCount=0 Then MsgBox ""No PCs marked Yes in CONFIG."",vbExclamation: GoTo Cleanup: End If" & vbCrLf
    s = s & "    Dim Xpc() As Double: ReDim Xpc(1 To g_n, 1 To pcCount)" & vbCrLf
    s = s & "    Dim ci As Long: Dim i As Long" & vbCrLf
    s = s & "    For j=1 To g_p: If inclPC(j) Then: ci=ci+1: For i=1 To g_n: Xpc(i,ci)=g_scores(i,j): Next i: End If: Next j" & vbCrLf
    s = s & "    Dim pcDates() As Variant: pcDates=g_dates" & vbCrLf
    s = s & "    OLSCore Xpc,g_Y,pcDates,g_n,pcCount,""PCA_REG"",True" & vbCrLf
    s = s & "    Dim beta() As Double: ReDim beta(1 To pcCount)" & vbCrLf
    s = s & "    Dim baseRow As Long: baseRow=15" & vbCrLf
    s = s & "    Dim wsC2 As Worksheet: Set wsC2=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim incInt As Boolean: incInt=(UCase(Trim(wsC2.Range(""INCLUDE_INT"").Value))=""YES"")" & vbCrLf
    s = s & "    If incInt Then baseRow=baseRow+1" & vbCrLf
    s = s & "    ci=0: For j=1 To g_p: If inclPC(j) Then: ci=ci+1: beta(ci)=wsR.Cells(baseRow+ci-1,3).Value: End If: Next j" & vbCrLf
    s = s & "    Dim fc() As Double: ReDim fc(1 To g_p): Dim totSq As Double" & vbCrLf
    s = s & "    For i=1 To g_p: Dim contrib As Double: contrib=0: ci=0" & vbCrLf
    s = s & "        For j=1 To g_p: If inclPC(j) Then: ci=ci+1: contrib=contrib+g_loadings(i,j)*beta(ci): End If: Next j" & vbCrLf
    s = s & "        fc(i)=contrib: totSq=totSq+contrib^2: Next i" & vbCrLf
    s = s & "    For i=1 To g_p: wsR.Cells(28+i,3)=fc(i): wsR.Cells(28+i,3).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        If totSq>0 Then wsR.Cells(28+i,4)=fc(i)^2/totSq: wsR.Cells(28+i,4).NumberFormat=""0.00%"": End If: Next i" & vbCrLf
    s = s & "    Application.Calculation=xlCalculationAutomatic: Application.ScreenUpdating=True" & vbCrLf
    s = s & "    MsgBox ""PCA Regression complete."",vbInformation,""PCA Analytics"": Exit Sub" & vbCrLf
    s = s & "Cleanup: Application.Calculation=xlCalculationAutomatic: Application.ScreenUpdating=True" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub OLSCore(X() As Double, Y() As Double, dates() As Variant, n As Long, p As Long, shNm As String, isPCA As Boolean)" & vbCrLf
    s = s & "    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets(shNm)" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim incInt As Boolean: incInt=(UCase(Trim(wsC.Range(""INCLUDE_INT"").Value))=""YES"")" & vbCrLf
    s = s & "    Dim confLvl As Double: confLvl=CDbl(wsC.Range(""CONF_LEVEL"").Value)" & vbCrLf
    s = s & "    Dim alpha As Double: alpha=1-confLvl" & vbCrLf
    s = s & "    Dim k As Long: k=p+IIf(incInt,1,0)" & vbCrLf
    s = s & "    Dim Xd() As Double: ReDim Xd(1 To n, 1 To k)" & vbCrLf
    s = s & "    Dim i As Long, j As Long" & vbCrLf
    s = s & "    If incInt Then: For i=1 To n: Xd(i,1)=1: Next i: For j=1 To p: For i=1 To n: Xd(i,j+1)=X(i,j): Next i: Next j" & vbCrLf
    s = s & "    Else: For i=1 To n: For j=1 To p: Xd(i,j)=X(i,j): Next j: Next i: End If" & vbCrLf
    s = s & "    Dim XtX() As Double: XtX=MatMul(MatTrans(Xd),Xd)" & vbCrLf
    s = s & "    Dim Xty() As Double: ReDim Xty(1 To k, 1 To 1)" & vbCrLf
    s = s & "    Dim sm As Double" & vbCrLf
    s = s & "    For j=1 To k: sm=0: For i=1 To n: sm=sm+Xd(i,j)*Y(i): Next i: Xty(j,1)=sm: Next j" & vbCrLf
    s = s & "    Dim XtXinv() As Double" & vbCrLf
    s = s & "    If Not MatInv(XtX,k,XtXinv) Then MsgBox ""X'X singular — multicollinearity."",vbCritical: Exit Sub: End If" & vbCrLf
    s = s & "    Dim beta() As Double: beta=MatMul(XtXinv,Xty)" & vbCrLf
    s = s & "    Dim yHat() As Double: ReDim yHat(1 To n): Dim resid() As Double: ReDim resid(1 To n)" & vbCrLf
    s = s & "    Dim yMean As Double: For i=1 To n: yMean=yMean+Y(i): Next i: yMean=yMean/n" & vbCrLf
    s = s & "    Dim SSR As Double, SST As Double, SSE As Double" & vbCrLf
    s = s & "    For i=1 To n: yHat(i)=0: For j=1 To k: yHat(i)=yHat(i)+Xd(i,j)*beta(j,1): Next j" & vbCrLf
    s = s & "        resid(i)=Y(i)-yHat(i): SSE=SSE+resid(i)^2: SSR=SSR+(yHat(i)-yMean)^2: SST=SST+(Y(i)-yMean)^2: Next i" & vbCrLf
    s = s & "    Dim dfE As Long: dfE=n-k" & vbCrLf
    s = s & "    Dim s2 As Double: s2=IIf(dfE>0,SSE/dfE,0)" & vbCrLf
    s = s & "    Dim tCrit As Double: tCrit=TInvCrit(alpha,dfE)" & vbCrLf
    s = s & "    Dim Rsq As Double: Rsq=IIf(SST>0,1-SSE/SST,0)" & vbCrLf
    s = s & "    Dim AdjR As Double: AdjR=IIf(dfE>0,1-(1-Rsq)*(n-1)/dfE,0)" & vbCrLf
    s = s & "    Dim dfR As Long: dfR=k-IIf(incInt,1,0)" & vbCrLf
    s = s & "    Dim Fst As Double: If dfR>0 And s2>0 Then Fst=(SSR/dfR)/s2" & vbCrLf
    s = s & "    Dim pF As Double: On Error Resume Next: pF=Application.WorksheetFunction.FDist(Fst,dfR,dfE): On Error GoTo 0" & vbCrLf
    s = s & "    Dim logL As Double: If s2>0 Then logL=-n/2*Log(6.28318*s2)-SSE/(2*s2)" & vbCrLf
    s = s & "    Dim AIC As Double: AIC=-2*logL+2*k" & vbCrLf
    s = s & "    Dim BIC As Double: BIC=-2*logL+Log(n)*k" & vbCrLf
    s = s & "    Dim depNm As String: depNm=CStr(wsC.Cells(4,2).Value)" & vbCrLf
    s = s & "    If isPCA Then" & vbCrLf
    s = s & "        ws.Cells(5,2)=depNm: ws.Cells(6,2)=p: ws.Cells(7,2)=n" & vbCrLf
    s = s & "        ws.Cells(8,2)=Rsq:  ws.Cells(8,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        ws.Cells(9,2)=AdjR: ws.Cells(9,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        ws.Cells(10,2)=Fst: ws.Cells(10,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        ws.Cells(11,2)=pF:  ws.Cells(11,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ws.Cells(5,2)=depNm: ws.Cells(6,2)=n" & vbCrLf
    s = s & "        ws.Cells(7,2)=Rsq:  ws.Cells(7,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        ws.Cells(8,2)=AdjR: ws.Cells(8,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        ws.Cells(9,2)=Fst:  ws.Cells(9,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        ws.Cells(10,2)=pF:  ws.Cells(10,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "        ws.Cells(11,2)=AIC: ws.Cells(11,2).NumberFormat=""0.00""" & vbCrLf
    s = s & "        ws.Cells(12,2)=BIC: ws.Cells(12,2).NumberFormat=""0.00""" & vbCrLf
    s = s & "        ws.Cells(16,2)=SSR: ws.Cells(16,3)=dfR: ws.Cells(16,4)=IIf(dfR>0,SSR/dfR,0): ws.Cells(16,5)=Fst: ws.Cells(16,6)=pF" & vbCrLf
    s = s & "        ws.Cells(17,2)=SSE: ws.Cells(17,3)=dfE: ws.Cells(17,4)=s2" & vbCrLf
    s = s & "        ws.Cells(18,2)=SST: ws.Cells(18,3)=n-1" & vbCrLf
    s = s & "        Dim at As Long: For at=16 To 18: Dim ac As Long: For ac=2 To 6: ws.Cells(at,ac).NumberFormat=""0.0000"": Next ac: Next at" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    Dim bRow As Long: bRow=IIf(isPCA,15,22)" & vbCrLf
    s = s & "    Dim jOff As Long: jOff=IIf(incInt,2,1)" & vbCrLf
    s = s & "    If incInt Then" & vbCrLf
    s = s & "        WriteCoeffRow ws,bRow,""Intercept"",beta(1,1),s2,XtXinv(1,1),tCrit,dfE: bRow=bRow+1" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    For j=1 To p" & vbCrLf
    s = s & "        Dim vNm As String" & vbCrLf
    s = s & "        If isPCA Then vNm=""PC"" & j Else vNm=CStr(wsC.Cells(14+j,3).Value)" & vbCrLf
    s = s & "        If Len(vNm)=0 Then vNm=""Factor "" & j" & vbCrLf
    s = s & "        WriteCoeffRow ws,bRow,vNm,beta(jOff+j-1,1),s2,XtXinv(jOff+j-1,jOff+j-1),tCrit,dfE" & vbCrLf
    s = s & "        bRow=bRow+1" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "    If Not isPCA Then" & vbCrLf
    s = s & "        For i=1 To n" & vbCrLf
    s = s & "            ws.Cells(36+i,1)=dates(i): ws.Cells(36+i,1).NumberFormat=""YYYY-MM-DD""" & vbCrLf
    s = s & "            ws.Cells(36+i,2)=Y(i): ws.Cells(36+i,3)=yHat(i): ws.Cells(36+i,4)=resid(i)" & vbCrLf
    s = s & "            ws.Cells(36+i,5)=IIf(Sqr(s2)>0,resid(i)/Sqr(s2),0)" & vbCrLf
    s = s & "            Dim cc As Long: For cc=2 To 5: ws.Cells(36+i,cc).NumberFormat=""0.0000"": Next cc" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub WriteCoeffRow(ws As Worksheet, r As Long, nm As String, b As Double, s2 As Double, diagXtXinv As Double, tCrit As Double, dfE As Long)" & vbCrLf
    s = s & "    Dim se As Double: se=Sqr(IIf(s2*diagXtXinv>=0,s2*diagXtXinv,0))" & vbCrLf
    s = s & "    Dim t As Double:  t=IIf(se>0,b/se,0)" & vbCrLf
    s = s & "    Dim p As Double:  p=TDistPValue(t,dfE)" & vbCrLf
    s = s & "    ws.Cells(r,1)=nm" & vbCrLf
    s = s & "    ws.Cells(r,2)=b:        ws.Cells(r,2).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    ws.Cells(r,3)=se:       ws.Cells(r,3).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    ws.Cells(r,4)=t:        ws.Cells(r,4).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    ws.Cells(r,5)=p:        ws.Cells(r,5).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    ws.Cells(r,6)=b-tCrit*se: ws.Cells(r,6).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    ws.Cells(r,7)=b+tCrit*se: ws.Cells(r,7).NumberFormat=""0.0000""" & vbCrLf
    s = s & "    Dim sig As String: If p<0.01 Then sig=""***"" ElseIf p<0.05 Then sig=""**"" ElseIf p<0.1 Then sig=""*"" Else sig="""": End If" & vbCrLf
    s = s & "    ws.Cells(r,8)=sig" & vbCrLf
    s = s & "    With ws.Cells(r,5).Font: .Bold=True" & vbCrLf
    s = s & "        If p<0.01 Then .Color=RGB(0,128,0) ElseIf p<0.05 Then .Color=RGB(0,176,80) ElseIf p<0.1 Then .Color=RGB(196,139,0) Else .Color=RGB(128,128,128): End If" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf

    c.CodeModule.AddFromString s
End Sub

Private Sub CreateModCharts(VBProj As Object)
    Dim c As Object
    Set c = VBProj.VBComponents.Add(1)
    c.Name = "modCharts"
    Dim s As String

    s = "Attribute VB_Name = ""modCharts""" & vbCrLf
    s = s & "Option Explicit" & vbCrLf & vbCrLf

    s = s & "Public Sub BuildPCACharts()" & vbCrLf
    s = s & "    ClearCharts" & vbCrLf
    s = s & "    BuildScreePlot: BuildCumVariancePlot: BuildLoadingsChart: BuildPCScoresChart" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub BuildRegressionCharts(): BuildResidualPlot: End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub ClearCharts()" & vbCrLf
    s = s & "    Dim ws As Worksheet: Set ws=ThisWorkbook.Sheets(""CHARTS"")" & vbCrLf
    s = s & "    Dim obj As ChartObject: For Each obj In ws.ChartObjects: obj.Delete: Next obj" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub BuildScreePlot()" & vbCrLf
    s = s & "    Dim wsP As Worksheet: Set wsP=ThisWorkbook.Sheets(""PCA"")" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim wsO As Worksheet: Set wsO=ThisWorkbook.Sheets(""CHARTS"")" & vbCrLf
    s = s & "    Dim nPC As Long: nPC=CLng(wsC.Range(""NUM_PCS"").Value)" & vbCrLf
    s = s & "    Dim tc As Long: tc=20: Dim j As Long" & vbCrLf
    s = s & "    For j=1 To nPC" & vbCrLf
    s = s & "        wsP.Cells(1,tc+j)=wsP.Cells(31+j,1).Value" & vbCrLf
    s = s & "        wsP.Cells(2,tc+j)=wsP.Cells(31+j,2).Value" & vbCrLf
    s = s & "        wsP.Cells(3,tc+j)=wsP.Cells(31+j,4).Value" & vbCrLf
    s = s & "    Next j" & vbCrLf
    s = s & "    Dim co As ChartObject: Set co=wsO.ChartObjects.Add(10,25,600,300)" & vbCrLf
    s = s & "    With co.Chart" & vbCrLf
    s = s & "        .ChartType=xlColumnClustered" & vbCrLf
    s = s & "        .HasTitle=True: .ChartTitle.Text=""Scree Plot  |  Red dashed = 80% cumulative threshold""" & vbCrLf
    s = s & "        .SeriesCollection.NewSeries" & vbCrLf
    s = s & "        With .SeriesCollection(1): .Name=""Eigenvalue""" & vbCrLf
    s = s & "            .Values=wsP.Range(wsP.Cells(2,tc+1),wsP.Cells(2,tc+nPC))" & vbCrLf
    s = s & "            .XValues=wsP.Range(wsP.Cells(1,tc+1),wsP.Cells(1,tc+nPC))" & vbCrLf
    s = s & "            .Interior.Color=RGB(46,117,182): .Border.Color=RGB(31,56,100)" & vbCrLf
    s = s & "        End With" & vbCrLf
    s = s & "        .SeriesCollection.NewSeries" & vbCrLf
    s = s & "        With .SeriesCollection(2): .Name=""Cumulative %""" & vbCrLf
    s = s & "            .Values=wsP.Range(wsP.Cells(3,tc+1),wsP.Cells(3,tc+nPC))" & vbCrLf
    s = s & "            .XValues=wsP.Range(wsP.Cells(1,tc+1),wsP.Cells(1,tc+nPC))" & vbCrLf
    s = s & "            .ChartType=xlLine: .AxisGroup=xlSecondary" & vbCrLf
    s = s & "            .Format.Line.ForeColor.RGB=RGB(201,168,76): .Format.Line.Weight=2.5" & vbCrLf
    s = s & "            .MarkerStyle=xlMarkerStyleCircle: .MarkerSize=5" & vbCrLf
    s = s & "        End With" & vbCrLf
    s = s & "        .SeriesCollection.NewSeries" & vbCrLf
    s = s & "        With .SeriesCollection(3): .Name=""80% Line""" & vbCrLf
    s = s & "            Dim v80() As Double: ReDim v80(1 To nPC): Dim k As Long: For k=1 To nPC: v80(k)=0.8: Next k" & vbCrLf
    s = s & "            .Values=v80: .ChartType=xlLine: .AxisGroup=xlSecondary" & vbCrLf
    s = s & "            .Format.Line.ForeColor.RGB=RGB(255,0,0): .Format.Line.DashStyle=msoLineDash" & vbCrLf
    s = s & "            .Format.Line.Weight=1.5: .MarkerStyle=xlMarkerStyleNone" & vbCrLf
    s = s & "        End With" & vbCrLf
    s = s & "        With .Axes(xlValue,xlPrimary): .HasTitle=True: .AxisTitle.Text=""Eigenvalue"": End With" & vbCrLf
    s = s & "        With .Axes(xlValue,xlSecondary): .HasTitle=True: .AxisTitle.Text=""Cumulative %""" & vbCrLf
    s = s & "            .MinimumScale=0: .MaximumScale=1: .NumberFormat=""0%"": End With" & vbCrLf
    s = s & "        .HasLegend=True: .Legend.Position=xlLegendPositionBottom" & vbCrLf
    s = s & "        StyleChart co.Chart" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub BuildCumVariancePlot()" & vbCrLf
    s = s & "    Dim wsP As Worksheet: Set wsP=ThisWorkbook.Sheets(""PCA"")" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim wsO As Worksheet: Set wsO=ThisWorkbook.Sheets(""CHARTS"")" & vbCrLf
    s = s & "    Dim nPC As Long: nPC=CLng(wsC.Range(""NUM_PCS"").Value): Dim j As Long" & vbCrLf
    s = s & "    Dim tc As Long: tc=35" & vbCrLf
    s = s & "    For j=1 To nPC: wsP.Cells(1,tc+j)=wsP.Cells(31+j,1).Value: wsP.Cells(2,tc+j)=wsP.Cells(31+j,4).Value: Next j" & vbCrLf
    s = s & "    Dim co As ChartObject: Set co=wsO.ChartObjects.Add(10,350,390,280)" & vbCrLf
    s = s & "    With co.Chart" & vbCrLf
    s = s & "        .ChartType=xlArea: .HasTitle=True: .ChartTitle.Text=""Cumulative Variance Explained""" & vbCrLf
    s = s & "        .SeriesCollection.NewSeries" & vbCrLf
    s = s & "        With .SeriesCollection(1): .Name=""Cum Var""" & vbCrLf
    s = s & "            .Values=wsP.Range(wsP.Cells(2,tc+1),wsP.Cells(2,tc+nPC))" & vbCrLf
    s = s & "            .XValues=wsP.Range(wsP.Cells(1,tc+1),wsP.Cells(1,tc+nPC))" & vbCrLf
    s = s & "            .Format.Fill.ForeColor.RGB=RGB(46,117,182)" & vbCrLf
    s = s & "        End With" & vbCrLf
    s = s & "        With .Axes(xlValue): .MinimumScale=0: .MaximumScale=1: .NumberFormat=""0%"": End With" & vbCrLf
    s = s & "        .HasLegend=False: StyleChart co.Chart" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub BuildLoadingsChart()" & vbCrLf
    s = s & "    Dim wsP As Worksheet: Set wsP=ThisWorkbook.Sheets(""PCA"")" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim wsO As Worksheet: Set wsO=ThisWorkbook.Sheets(""CHARTS"")" & vbCrLf
    s = s & "    Dim nPC As Long: nPC=CLng(wsC.Range(""NUM_PCS"").Value)" & vbCrLf
    s = s & "    Dim p As Long: p=g_p: If p=0 Then Exit Sub" & vbCrLf
    s = s & "    Dim tr As Long: tr=70: Dim j As Long, i As Long" & vbCrLf
    s = s & "    For j=1 To nPC: wsP.Cells(tr,j+1)=wsP.Cells(31+j,1).Value: Next j" & vbCrLf
    s = s & "    For i=1 To p: wsP.Cells(tr+i,1)=wsC.Cells(14+i,3).Value" & vbCrLf
    s = s & "        For j=1 To nPC: wsP.Cells(tr+i,j+1)=wsP.Cells(44+i,j+1).Value: Next j: Next i" & vbCrLf
    s = s & "    Dim co As ChartObject: Set co=wsO.ChartObjects.Add(415,350,600,280)" & vbCrLf
    s = s & "    Dim clrs As Variant: clrs=Array(RGB(46,117,182),RGB(112,173,71),RGB(201,168,76),RGB(237,125,49),RGB(112,48,160))" & vbCrLf
    s = s & "    With co.Chart" & vbCrLf
    s = s & "        .ChartType=xlBarClustered: .HasTitle=True: .ChartTitle.Text=""Factor Loadings by PC  (bold = |loading| >= 0.5)""" & vbCrLf
    s = s & "        For j=1 To IIf(nPC>5,5,nPC)" & vbCrLf
    s = s & "            .SeriesCollection.NewSeries" & vbCrLf
    s = s & "            With .SeriesCollection(j): .Name=wsP.Cells(tr,j+1).Value" & vbCrLf
    s = s & "                .Values=wsP.Range(wsP.Cells(tr+1,j+1),wsP.Cells(tr+p,j+1))" & vbCrLf
    s = s & "                .XValues=wsP.Range(wsP.Cells(tr+1,1),wsP.Cells(tr+p,1))" & vbCrLf
    s = s & "                .Interior.Color=clrs(j-1)" & vbCrLf
    s = s & "            End With: Next j" & vbCrLf
    s = s & "        .HasLegend=True: .Legend.Position=xlLegendPositionRight: StyleChart co.Chart" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub BuildPCScoresChart()" & vbCrLf
    s = s & "    Dim wsP As Worksheet: Set wsP=ThisWorkbook.Sheets(""PCA"")" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim wsO As Worksheet: Set wsO=ThisWorkbook.Sheets(""CHARTS"")" & vbCrLf
    s = s & "    Dim nPC As Long: nPC=CLng(wsC.Range(""NUM_PCS"").Value): Dim n As Long: n=g_n" & vbCrLf
    s = s & "    If n=0 Then Exit Sub" & vbCrLf
    s = s & "    Dim co As ChartObject: Set co=wsO.ChartObjects.Add(10,655,600,280)" & vbCrLf
    s = s & "    Dim lclrs As Variant: lclrs=Array(RGB(46,117,182),RGB(201,168,76),RGB(112,173,71),RGB(237,125,49),RGB(112,48,160))" & vbCrLf
    s = s & "    With co.Chart" & vbCrLf
    s = s & "        .ChartType=xlLine: .HasTitle=True: .ChartTitle.Text=""PC Scores — Time Series""" & vbCrLf
    s = s & "        Dim j As Long" & vbCrLf
    s = s & "        For j=1 To IIf(nPC>5,5,nPC)" & vbCrLf
    s = s & "            .SeriesCollection.NewSeries" & vbCrLf
    s = s & "            With .SeriesCollection(j): .Name=wsP.Cells(31+j,1).Value" & vbCrLf
    s = s & "                .Values=wsP.Range(wsP.Cells(58,j+1),wsP.Cells(57+n,j+1))" & vbCrLf
    s = s & "                .XValues=wsP.Range(wsP.Cells(58,1),wsP.Cells(57+n,1))" & vbCrLf
    s = s & "                .Format.Line.ForeColor.RGB=lclrs(j-1): .Format.Line.Weight=1.5" & vbCrLf
    s = s & "                .MarkerStyle=xlMarkerStyleNone" & vbCrLf
    s = s & "            End With: Next j" & vbCrLf
    s = s & "        .HasLegend=True: .Legend.Position=xlLegendPositionBottom: StyleChart co.Chart" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub BuildResidualPlot()" & vbCrLf
    s = s & "    Dim wsR As Worksheet: Set wsR=ThisWorkbook.Sheets(""REGRESSION"")" & vbCrLf
    s = s & "    Dim wsO As Worksheet: Set wsO=ThisWorkbook.Sheets(""CHARTS"")" & vbCrLf
    s = s & "    Dim n As Long: n=g_n: If n=0 Then Exit Sub" & vbCrLf
    s = s & "    Dim co As ChartObject: Set co=wsO.ChartObjects.Add(415,655,600,280)" & vbCrLf
    s = s & "    With co.Chart: .ChartType=xlXYScatter: .HasTitle=True: .ChartTitle.Text=""Residuals vs Fitted""" & vbCrLf
    s = s & "        .SeriesCollection.NewSeries" & vbCrLf
    s = s & "        With .SeriesCollection(1): .Name=""Residuals""" & vbCrLf
    s = s & "            .XValues=wsR.Range(wsR.Cells(37,3),wsR.Cells(36+n,3))" & vbCrLf
    s = s & "            .Values=wsR.Range(wsR.Cells(37,4),wsR.Cells(36+n,4))" & vbCrLf
    s = s & "            .MarkerStyle=xlMarkerStyleCircle: .MarkerSize=4" & vbCrLf
    s = s & "            .MarkerForegroundColor=RGB(46,117,182): .Format.Line.Visible=msoFalse" & vbCrLf
    s = s & "        End With" & vbCrLf
    s = s & "        .HasLegend=False: StyleChart co.Chart" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub StyleChart(cht As Chart)" & vbCrLf
    s = s & "    With cht.PlotArea: .Interior.Color=RGB(242,242,242): .Border.Color=RGB(217,217,217): End With" & vbCrLf
    s = s & "    cht.ChartArea.Border.Color=RGB(217,217,217): cht.ChartTitle.Font.Name=""Calibri""" & vbCrLf
    s = s & "End Sub" & vbCrLf

    c.CodeModule.AddFromString s
End Sub

Private Sub CreateModMain(VBProj As Object)
    Dim c As Object
    Set c = VBProj.VBComponents.Add(1)
    c.Name = "modMain"
    Dim s As String

    s = "Attribute VB_Name = ""modMain""" & vbCrLf
    s = s & "Option Explicit" & vbCrLf & vbCrLf

    s = s & "Public Sub Auto_Open()" & vbCrLf
    s = s & "    Application.OnKey ""^+P"", ""RunPCA_Entry""" & vbCrLf
    s = s & "    Application.OnKey ""^+R"", ""RunRegression_Entry""" & vbCrLf
    s = s & "    Application.OnKey ""^+Q"", ""RunPCAReg_Entry""" & vbCrLf
    s = s & "    Application.OnKey ""^+A"", ""RunAll""" & vbCrLf
    s = s & "    Application.OnKey ""^+X"", ""ResetOutputs""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub Auto_Close()" & vbCrLf
    s = s & "    Application.OnKey ""^+P"": Application.OnKey ""^+R""" & vbCrLf
    s = s & "    Application.OnKey ""^+Q"": Application.OnKey ""^+A"": Application.OnKey ""^+X""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub RunPCA_Entry(): RunPCA: End Sub" & vbCrLf
    s = s & "Public Sub RunRegression_Entry(): RunRegression: End Sub" & vbCrLf
    s = s & "Public Sub RunPCAReg_Entry(): RunPCARegression: End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub RunAll()" & vbCrLf
    s = s & "    Dim t0 As Single: t0=Timer: Application.ScreenUpdating=False" & vbCrLf
    s = s & "    RunPCA: RunRegression: RunPCARegression: BuildRegressionCharts" & vbCrLf
    s = s & "    Application.ScreenUpdating=True" & vbCrLf
    s = s & "    MsgBox ""All analyses complete in "" & Format(Timer-t0,""0.0"") & ""s."" & vbCrLf & vbCrLf & ""PCA sheet: eigenvalues, loadings, scores"" & vbCrLf & ""REGRESSION: OLS coefficients, ANOVA, residuals"" & vbCrLf & ""PCA_REG: PC regression + factor contributions"" & vbCrLf & ""CHARTS: all charts updated"",vbInformation,""PCA Analytics""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub ResetOutputs()" & vbCrLf
    s = s & "    If MsgBox(""Clear all computed outputs (PCA/REGRESSION/PCA_REG/CHARTS)?"",vbYesNo+vbQuestion,""Reset"")=vbNo Then Exit Sub" & vbCrLf
    s = s & "    Dim ws As Worksheet" & vbCrLf
    s = s & "    Set ws=ThisWorkbook.Sheets(""PCA"")" & vbCrLf
    s = s & "    ClrRange ws,6,15,2,6: ClrRange ws,19,28,2,11: ClrRange ws,32,41,2,6" & vbCrLf
    s = s & "    ClrRange ws,45,54,2,11: ClrRange ws,58,300,1,11: ClrRange ws,1,10,20,50: ClrRange ws,70,90,1,15" & vbCrLf
    s = s & "    Set ws=ThisWorkbook.Sheets(""REGRESSION"")" & vbCrLf
    s = s & "    ClrRange ws,5,12,2,2: ClrRange ws,16,18,2,6: ClrRange ws,22,33,1,8: ClrRange ws,37,300,1,5" & vbCrLf
    s = s & "    Set ws=ThisWorkbook.Sheets(""PCA_REG"")" & vbCrLf
    s = s & "    ClrRange ws,5,11,2,2: ClrRange ws,15,26,1,8: ClrRange ws,29,38,3,4" & vbCrLf
    s = s & "    Dim obj As ChartObject: For Each obj In ThisWorkbook.Sheets(""CHARTS"").ChartObjects: obj.Delete: Next obj" & vbCrLf
    s = s & "    g_n=0: g_p=0" & vbCrLf
    s = s & "    MsgBox ""Outputs cleared."",vbInformation,""PCA Analytics""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Private Sub ClrRange(ws As Worksheet, r1 As Long, r2 As Long, c1 As Long, c2 As Long)" & vbCrLf
    s = s & "    Dim r As Long, c As Long" & vbCrLf
    s = s & "    For r=r1 To r2: For c=c1 To c2" & vbCrLf
    s = s & "        If Not ws.Cells(r,c).HasFormula Then ws.Cells(r,c).ClearContents" & vbCrLf
    s = s & "    Next c: Next r" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf

    s = s & "Public Sub ShowDiagnostics()" & vbCrLf
    s = s & "    Dim wsC As Worksheet: Set wsC=ThisWorkbook.Sheets(""CONFIG"")" & vbCrLf
    s = s & "    Dim msg As String" & vbCrLf
    s = s & "    msg=""Data rows: "" & CLng(wsC.Range(""DATA_END_ROW"").Value)-CLng(wsC.Range(""DATA_START_ROW"").Value)+1 & vbCrLf" & vbCrLf
    s = s & "    msg=msg & ""Matrix type: "" & wsC.Range(""MATRIX_TYPE"").Value & vbCrLf" & vbCrLf
    s = s & "    msg=msg & ""PCs retained: "" & wsC.Range(""NUM_PCS"").Value & vbCrLf" & vbCrLf
    s = s & "    msg=msg & ""Last run n="" & g_n & ""  p="" & g_p & vbCrLf" & vbCrLf
    s = s & "    If g_n>0 Then" & vbCrLf
    s = s & "        Dim tot As Double: Dim j As Long: For j=1 To g_p: tot=tot+g_eigenVals(j): Next j" & vbCrLf
    s = s & "        Dim cum As Double" & vbCrLf
    s = s & "        msg=msg & vbCrLf & ""Top eigenvalues:"" & vbCrLf" & vbCrLf
    s = s & "        For j=1 To IIf(g_p>5,5,g_p): cum=cum+IIf(tot>0,g_eigenVals(j)/tot,0)" & vbCrLf
    s = s & "            msg=msg & ""  PC"" & j & "": "" & Format(g_eigenVals(j),""0.0000"") & ""  (cum="" & Format(cum*100,""0.1"") & ""%)"" & vbCrLf" & vbCrLf
    s = s & "        Next j" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    MsgBox msg,vbInformation,""Diagnostics""" & vbCrLf
    s = s & "End Sub" & vbCrLf

    c.CodeModule.AddFromString s
End Sub
