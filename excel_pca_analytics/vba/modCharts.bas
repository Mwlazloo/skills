Attribute VB_Name = "modCharts"
'=============================================================================
' modCharts — Chart creation for PCA Analytics
'
' Charts produced:
'   1. Scree Plot      — bar (eigenvalues) + line (cumulative %) combo
'   2. Cumulative Variance Explained — area chart
'   3. Factor Loadings Heatmap — clustered bar per PC
'   4. PC Scores Time Series   — line chart
'   5. Residual Plot           — scatter (fitted vs residual)
'=============================================================================
Option Explicit

Private Const CHARTS_SHEET As String = "CHARTS"

' ── Main entry called by RunPCA ───────────────────────────────────────────────
Public Sub BuildPCACharts()
    ClearCharts
    BuildScreePlot
    BuildCumVariancePlot
    BuildLoadingsChart
    BuildPCScoresChart
End Sub

' ── Called by RunRegression ───────────────────────────────────────────────────
Public Sub BuildRegressionCharts()
    BuildResidualPlot
End Sub

' ── Remove all existing charts from CHARTS sheet ─────────────────────────────
Private Sub ClearCharts()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(CHARTS_SHEET)
    Dim obj As ChartObject
    For Each obj In ws.ChartObjects
        obj.Delete
    Next obj
End Sub

'=============================================================================
' 1. Scree Plot — combo: bars for eigenvalues, secondary-axis line for cum%
'=============================================================================
Private Sub BuildScreePlot()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PCA")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")
    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Sheets(CHARTS_SHEET)
    Dim numPCs As Long: numPCs = CLng(wsC.Range("NUM_PCS").Value)

    ' Write temp data for chart (col S of PCA sheet to avoid layout disruption)
    Dim tempCol As Long: tempCol = 20   ' column T
    wsP.Cells(1, tempCol).Value = "PC Label"
    wsP.Cells(2, tempCol).Value = "Eigenvalue"
    wsP.Cells(3, tempCol).Value = "Cum %"
    Dim j As Long
    For j = 1 To numPCs
        wsP.Cells(1, tempCol + j).Value = wsP.Cells(31 + j, 1).Value   ' PC label
        wsP.Cells(2, tempCol + j).Value = wsP.Cells(31 + j, 2).Value   ' eigenvalue
        wsP.Cells(3, tempCol + j).Value = wsP.Cells(31 + j, 4).Value   ' cum %
    Next j

    ' Create chart
    Dim co As ChartObject
    Set co = wsOut.ChartObjects.Add(Left:=10, Top:=25, Width:=580, Height:=290)
    With co.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Scree Plot — Eigenvalues & Cumulative Variance"
        .ChartTitle.Font.Size = 11: .ChartTitle.Font.Bold = True

        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Eigenvalue"
            .Values = wsP.Range(wsP.Cells(2, tempCol + 1), wsP.Cells(2, tempCol + numPCs))
            .XValues = wsP.Range(wsP.Cells(1, tempCol + 1), wsP.Cells(1, tempCol + numPCs))
            .ChartType = xlColumnClustered
            .Interior.Color = RGB(46, 117, 182)
            .Border.Color = RGB(31, 56, 100)
        End With

        .SeriesCollection.NewSeries
        With .SeriesCollection(2)
            .Name = "Cumulative %"
            .Values = wsP.Range(wsP.Cells(3, tempCol + 1), wsP.Cells(3, tempCol + numPCs))
            .XValues = wsP.Range(wsP.Cells(1, tempCol + 1), wsP.Cells(1, tempCol + numPCs))
            .ChartType = xlLine
            .AxisGroup = xlSecondary
            .Interior.Color = RGB(201, 168, 76)
            .Format.Line.ForeColor.RGB = RGB(201, 168, 76)
            .Format.Line.Weight = 2.5
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
        End With

        ' Primary axis — eigenvalues
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Eigenvalue"
            .AxisTitle.Font.Size = 9
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
        End With

        ' Secondary axis — cumulative %
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "Cumulative Variance Explained"
            .AxisTitle.Font.Size = 9
            .MinimumScale = 0
            .MaximumScale = 1
            .NumberFormat = "0%"
        End With

        ' Add Kaiser criterion line (eigenvalue = 1) as error bar workaround
        ' Simplest approach: annotate the chart title
        .ChartTitle.Text = "Scree Plot  |  Dashed line = 80% threshold"

        ' Reference line at 80% on secondary axis via a dummy series
        .SeriesCollection.NewSeries
        With .SeriesCollection(3)
            .Name = "80% Threshold"
            Dim vals80() As Double: ReDim vals80(1 To numPCs)
            Dim k As Long
            For k = 1 To numPCs: vals80(k) = 0.8: Next k
            .Values = vals80
            .ChartType = xlLine
            .AxisGroup = xlSecondary
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 1.5
            .MarkerStyle = xlMarkerStyleNone
        End With

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        StylePlotArea co.Chart
    End With
End Sub

'=============================================================================
' 2. Cumulative Variance Explained — area chart
'=============================================================================
Private Sub BuildCumVariancePlot()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PCA")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")
    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Sheets(CHARTS_SHEET)
    Dim numPCs As Long: numPCs = CLng(wsC.Range("NUM_PCS").Value)

    Dim tempCol As Long: tempCol = 35   ' col AJ on PCA sheet
    For j = 1 To numPCs
        wsP.Cells(1, tempCol + j).Value = wsP.Cells(31 + j, 1).Value
        wsP.Cells(2, tempCol + j).Value = wsP.Cells(31 + j, 4).Value
        wsP.Cells(3, tempCol + j).Value = wsP.Cells(31 + j, 3).Value
    Next j

    Dim co As ChartObject
    Set co = wsOut.ChartObjects.Add(Left:=10, Top:=340, Width:=380, Height:=280)
    With co.Chart
        .ChartType = xlAreaStacked
        .HasTitle = True
        .ChartTitle.Text = "Cumulative Variance Explained"
        .ChartTitle.Font.Size = 11

        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Cum. Variance"
            .Values = wsP.Range(wsP.Cells(2, tempCol + 1), wsP.Cells(2, tempCol + numPCs))
            .XValues = wsP.Range(wsP.Cells(1, tempCol + 1), wsP.Cells(1, tempCol + numPCs))
            .Format.Fill.ForeColor.RGB = RGB(46, 117, 182)
            .Format.Fill.Transparency = 0.3
        End With

        With .Axes(xlValue)
            .MinimumScale = 0: .MaximumScale = 1
            .NumberFormat = "0%"
        End With

        .HasLegend = False
        StylePlotArea co.Chart
    End With
End Sub

'=============================================================================
' 3. Factor Loadings Bar Chart — one cluster per PC
'    Shows how much each factor loads on each PC
'=============================================================================
Private Sub BuildLoadingsChart()
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Sheets("PCA")
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")
    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Sheets(CHARTS_SHEET)
    Dim numPCs As Long: numPCs = CLng(wsC.Range("NUM_PCS").Value)
    Dim p As Long: p = g_p
    If p = 0 Then Exit Sub

    ' Write loadings data to temp area
    Dim tempRow As Long: tempRow = 70
    wsP.Cells(tempRow, 1).Value = "Factor"
    Dim j As Long
    For j = 1 To numPCs
        wsP.Cells(tempRow, j + 1).Value = wsP.Cells(31 + j, 1).Value   ' PC label
    Next j
    Dim i As Long
    For i = 1 To p
        wsP.Cells(tempRow + i, 1).Value = wsC.Cells(14 + i, 3).Value   ' factor display name
        For j = 1 To numPCs
            wsP.Cells(tempRow + i, j + 1).Value = wsP.Cells(44 + i, j + 1).Value
        Next j
    Next i

    Dim co As ChartObject
    Set co = wsOut.ChartObjects.Add(Left:=400, Top:=340, Width:=580, Height:=280)
    With co.Chart
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Factor Loadings by Principal Component"
        .ChartTitle.Font.Size = 11

        ' Add one series per PC
        Dim clrs As Variant
        clrs = Array(RGB(46, 117, 182), RGB(112, 173, 71), RGB(201, 168, 76), _
                     RGB(237, 125, 49), RGB(112, 48, 160), RGB(0, 176, 240), _
                     RGB(255, 0, 0), RGB(68, 114, 196), RGB(32, 178, 170), RGB(139, 139, 0))

        For j = 1 To numPCs
            .SeriesCollection.NewSeries
            With .SeriesCollection(j)
                .Name = wsP.Cells(tempRow, j + 1).Value
                .Values = wsP.Range(wsP.Cells(tempRow + 1, j + 1), wsP.Cells(tempRow + p, j + 1))
                .XValues = wsP.Range(wsP.Cells(tempRow + 1, 1), wsP.Cells(tempRow + p, 1))
                .Interior.Color = clrs((j - 1) Mod 10)
            End With
        Next j

        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Loading"
            .AxisTitle.Font.Size = 9
        End With

        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        StylePlotArea co.Chart
    End With
End Sub

'=============================================================================
' 4. PC Scores Time Series — line chart, one line per retained PC
'=============================================================================
Private Sub BuildPCScoresChart()
    Dim wsP  As Worksheet: Set wsP = ThisWorkbook.Sheets("PCA")
    Dim wsC  As Worksheet: Set wsC = ThisWorkbook.Sheets("CONFIG")
    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Sheets(CHARTS_SHEET)
    Dim numPCs As Long: numPCs = CLng(wsC.Range("NUM_PCS").Value)
    Dim n As Long: n = g_n
    If n = 0 Then Exit Sub

    Dim co As ChartObject
    Set co = wsOut.ChartObjects.Add(Left:=600, Top:=340, Width:=580, Height:=280)
    With co.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "PC Scores — Time Series"
        .ChartTitle.Font.Size = 11

        Dim j As Long
        Dim lineClrs As Variant
        lineClrs = Array(RGB(46, 117, 182), RGB(201, 168, 76), RGB(112, 173, 71), _
                         RGB(237, 125, 49), RGB(112, 48, 160))

        For j = 1 To numPCs
            .SeriesCollection.NewSeries
            With .SeriesCollection(j)
                .Name = wsP.Cells(31 + j, 1).Value
                .Values = wsP.Range(wsP.Cells(58, j + 1), wsP.Cells(57 + n, j + 1))
                .XValues = wsP.Range(wsP.Cells(58, 1), wsP.Cells(57 + n, 1))
                .Format.Line.ForeColor.RGB = lineClrs((j - 1) Mod 5)
                .Format.Line.Weight = 1.25
                .MarkerStyle = xlMarkerStyleNone
            End With
        Next j

        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Score"
            .AxisTitle.Font.Size = 9
        End With

        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        StylePlotArea co.Chart
    End With
End Sub

'=============================================================================
' 5. Residual Diagnostic Plot — fitted vs residuals scatter
'=============================================================================
Public Sub BuildResidualPlot()
    Dim wsR  As Worksheet: Set wsR  = ThisWorkbook.Sheets("REGRESSION")
    Dim wsOut As Worksheet: Set wsOut = ThisWorkbook.Sheets(CHARTS_SHEET)
    Dim n As Long: n = g_n
    If n = 0 Then Exit Sub

    Dim co As ChartObject
    Set co = wsOut.ChartObjects.Add(Left:=10, Top:=680, Width:=450, Height:=280)
    With co.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Text = "Residuals vs Fitted Values"
        .ChartTitle.Font.Size = 11

        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Residuals"
            .XValues = wsR.Range(wsR.Cells(37, 3), wsR.Cells(36 + n, 3))   ' fitted
            .Values  = wsR.Range(wsR.Cells(37, 4), wsR.Cells(36 + n, 4))   ' residual
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 4
            .MarkerForegroundColor = RGB(46, 117, 182)
            .MarkerBackgroundColor = RGB(46, 117, 182)
            .Format.Line.Visible = msoFalse
        End With

        With .Axes(xlValue)
            .HasTitle = True: .AxisTitle.Text = "Residual": .AxisTitle.Font.Size = 9
        End With
        With .Axes(xlCategory)
            .HasTitle = True: .AxisTitle.Text = "Fitted": .AxisTitle.Font.Size = 9
        End With

        .HasLegend = False
        StylePlotArea co.Chart
    End With
End Sub

' ── Shared plot area styling ───────────────────────────────────────────────────
Private Sub StylePlotArea(cht As Chart)
    With cht.PlotArea
        .Interior.Color = RGB(242, 242, 242)
        .Border.Color = RGB(217, 217, 217)
    End With
    With cht
        .ChartArea.Border.Color = RGB(217, 217, 217)
        .ChartArea.Interior.Color = RGB(255, 255, 255)
        .ChartTitle.Font.Name = "Calibri"
        Dim ax As Axis
        For Each ax In .Axes
            ax.TickLabels.Font.Size = 8
            ax.TickLabels.Font.Name = "Calibri"
        Next ax
    End With
End Sub
