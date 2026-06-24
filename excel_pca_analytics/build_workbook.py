"""
build_workbook.py
Creates the PCA Analytics Excel workbook structure.
Run: pip install openpyxl && python build_workbook.py

After running:
  1. Open PCA_Analytics.xlsm in Excel
  2. Alt+F11 → File → Import File → import each vba/*.bas file
  3. Close VBA editor, save as .xlsm (macro-enabled)
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
import openpyxl.chart.series
from openpyxl.drawing.image import Image

# ── Color Palette ──────────────────────────────────────────────────────────────
C_NAVY      = "1F3864"   # Deep navy — primary headers
C_BLUE      = "2E75B6"   # Mid blue — section headers
C_LTBLUE    = "D6E4F0"   # Light blue — sub-headers
C_GOLD      = "C9A84C"   # Gold — accent / highlight
C_WHITE     = "FFFFFF"
C_LGRAY     = "F2F2F2"   # Light gray — alternating rows
C_MGRAY     = "D9D9D9"   # Medium gray — borders
C_GREEN     = "70AD47"   # Green — positive values
C_RED       = "FF0000"   # Red — errors / negatives
C_ORANGE    = "ED7D31"   # Orange — warnings

# ── Style Helpers ──────────────────────────────────────────────────────────────
def hdr_style(cell, text, bg=C_NAVY, fg=C_WHITE, bold=True, size=11, center=True):
    cell.value = text
    cell.font = Font(name="Calibri", bold=bold, color=fg, size=size)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(
        horizontal="center" if center else "left",
        vertical="center", wrap_text=True
    )

def section_hdr(cell, text):
    hdr_style(cell, text, bg=C_BLUE, size=10)

def sub_hdr(cell, text):
    hdr_style(cell, text, bg=C_LTBLUE, fg=C_NAVY, size=10)

def label(cell, text, bold=False):
    cell.value = text
    cell.font = Font(name="Calibri", bold=bold, size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center")

def val(cell, v, fmt=None):
    cell.value = v
    cell.font = Font(name="Calibri", size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        cell.number_format = fmt

def thin_border(ws, min_row, min_col, max_row, max_col, color=C_MGRAY):
    side = Side(style="thin", color=color)
    for row in ws.iter_rows(min_row=min_row, min_col=min_col,
                             max_row=max_row, max_col=max_col):
        for c in row:
            c.border = Border(left=side, right=side, top=side, bottom=side)

def merge_hdr(ws, cell_range, text, bg=C_NAVY, fg=C_WHITE, size=12):
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":")[0]]
    hdr_style(cell, text, bg=bg, fg=fg, size=size)

def set_col_width(ws, col_widths: dict):
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

def freeze(ws, cell):
    ws.freeze_panes = cell

# ── Workbook Setup ─────────────────────────────────────────────────────────────
wb = Workbook()
wb.remove(wb.active)   # remove default Sheet

# Order matters for tab display
sheets = {}
for name in ["DATA", "CONFIG", "PCA", "REGRESSION", "PCA_REG", "CHARTS"]:
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = (name == "DATA")   # grids only on DATA
    sheets[name] = ws

TAB_COLORS = {
    "DATA":       "2E75B6",
    "CONFIG":     "1F3864",
    "PCA":        "70AD47",
    "REGRESSION": "C9A84C",
    "PCA_REG":    "ED7D31",
    "CHARTS":     "7030A0",
}
for name, color in TAB_COLORS.items():
    sheets[name].sheet_properties.tabColor = color


# ══════════════════════════════════════════════════════════════════════════════
# 1. DATA SHEET
# ══════════════════════════════════════════════════════════════════════════════
ws = sheets["DATA"]
ws.sheet_view.showGridLines = True

# Title banner
ws.row_dimensions[1].height = 28
ws.merge_cells("A1:L1")
hdr_style(ws["A1"], "PCA ANALYTICS  ·  DATA INPUT", bg=C_NAVY, size=14)

# Instruction row
ws.row_dimensions[2].height = 18
ws.merge_cells("A2:L2")
ws["A2"].value = (
    "Replace placeholder formulas with CIQ PRO =CIQ() functions. "
    "Keep headers in Row 4. Data starts Row 5. Add/remove factor columns freely."
)
ws["A2"].font = Font(name="Calibri", italic=True, color=C_BLUE, size=9)
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
ws["A2"].fill = PatternFill("solid", fgColor="EBF3FB")

ws.row_dimensions[3].height = 6   # spacer

# Column headers — Row 4
ws.row_dimensions[4].height = 22
HEADERS = [
    "Date",
    "Dep Var Return",
    "Factor 1",
    "Factor 2",
    "Factor 3",
    "Factor 4",
    "Factor 5",
    "Factor 6",
    "Factor 7",
    "Factor 8",
    "Factor 9",
    "Factor 10",
]
for i, h in enumerate(HEADERS, start=1):
    cell = ws.cell(row=4, column=i)
    hdr_style(cell, h, bg=C_NAVY if i <= 2 else C_BLUE, size=10)

# Sample placeholder data rows (5–54 = 50 rows)
DATE_FMT = "YYYY-MM-DD"
NUM_FMT  = '0.0000'
for r in range(5, 55):
    ws.row_dimensions[r].height = 16
    # Date column
    c = ws.cell(row=r, column=1)
    c.value = f'=IFERROR(DATE(2020,1,{r-4}),"")' if r == 5 else f"=A{r-1}+7"
    c.number_format = DATE_FMT
    c.font = Font(name="Calibri", size=10)
    c.alignment = Alignment(horizontal="center")
    # Numeric placeholders — user replaces with =CIQ(...) formulas
    for col in range(2, len(HEADERS) + 1):
        c = ws.cell(row=r, column=col)
        c.value = None   # blank — user fills with CIQ formulas
        c.number_format = NUM_FMT
        c.font = Font(name="Calibri", size=10)
        c.alignment = Alignment(horizontal="right")
        # Alternating row shading
        if r % 2 == 0:
            c.fill = PatternFill("solid", fgColor=C_LGRAY)

# Thin borders on data range
thin_border(ws, 4, 1, 54, len(HEADERS))

# Freeze date + header
freeze(ws, "B5")

# Column widths
col_widths = {"A": 13, "B": 16}
for col_letter in [get_column_letter(i) for i in range(3, len(HEADERS) + 1)]:
    col_widths[col_letter] = 13
set_col_width(ws, col_widths)

# Named range: DATA_TABLE (VBA uses this)
# Note: Named ranges set after all sheets are built (see bottom)

# CIQ note box (columns M–P)
ws.merge_cells("N4:P4")
hdr_style(ws["N4"], "CIQ PRO Formula Reference", bg=C_GOLD, fg=C_NAVY, size=10)
ciq_notes = [
    ("Single value:", "=CIQ(\"TICKER\",\"IQ_TOTAL_RETURN\",\"IQ_FY0\")"),
    ("Time series:", "=CIQRANGE(\"TICKER\",\"IQ_TOTAL_RETURN\",start,end,\"IQ_MONTHLY\")"),
    ("Index return:", "=CIQ(\"SP500\",\"IQ_TOTAL_RETURN\",\"IQ_FY-1\")"),
    ("Tip:", "Paste Transpose for vertical time series"),
]
for i, (lbl_text, formula) in enumerate(ciq_notes, start=5):
    ws.cell(row=i, column=14).value = lbl_text
    ws.cell(row=i, column=14).font = Font(name="Calibri", bold=True, size=9, color=C_NAVY)
    ws.merge_cells(f"O{i}:P{i}")
    c = ws.cell(row=i, column=15)
    c.value = formula
    c.font = Font(name="Courier New", size=9, color=C_BLUE)
    c.alignment = Alignment(horizontal="left")

for col in ["N", "O", "P"]:
    ws.column_dimensions[col].width = 20


# ══════════════════════════════════════════════════════════════════════════════
# 2. CONFIG SHEET
# ══════════════════════════════════════════════════════════════════════════════
ws = sheets["CONFIG"]
ws.sheet_view.showGridLines = False

# Title
ws.row_dimensions[1].height = 30
ws.merge_cells("A1:F1")
hdr_style(ws["A1"], "CONFIGURATION & LABELS", bg=C_NAVY, size=14)

set_col_width(ws, {"A": 24, "B": 22, "C": 22, "D": 20, "E": 18, "F": 16})

# ── Section A: Analysis Settings ─────────────────────────────────────────────
ws.row_dimensions[3].height = 20
ws.merge_cells("A3:F3")
section_hdr(ws["A3"], "A.  ANALYSIS SETTINGS")

settings = [
    (4,  "Dependent Variable Name",  "Portfolio Return",    "Displayed in all outputs"),
    (5,  "Number of Factors (auto)", '=COUNTA(DATA!B4:K4)-1', "Auto-counted — do not edit"),
    (6,  "PCs to Retain",            5,                     "PCs shown in output / charts"),
    (7,  "Matrix Type",              "Correlation",         "Correlation or Covariance"),
    (8,  "Confidence Level",         0.95,                  "For regression CIs (e.g. 0.95)"),
    (9,  "Include Intercept",        "Yes",                 "Yes / No"),
    (10, "Data Start Row (DATA)",    5,                     "First data row on DATA sheet"),
    (11, "Data End Row (DATA)",      '=COUNTA(DATA!A:A)+3', "Last data row (auto)"),
]

for row, lbl_text, default, note in settings:
    ws.row_dimensions[row].height = 18
    sub_hdr(ws.cell(row=row, column=1), lbl_text)
    c = ws.cell(row=row, column=2)
    c.value = default
    c.font = Font(name="Calibri", bold=True, size=10, color=C_NAVY)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor="FFF2CC")   # yellow = editable
    c = ws.cell(row=row, column=3)
    c.value = note
    c.font = Font(name="Calibri", italic=True, size=9, color="595959")

# Dropdown validators
dv_matrix = DataValidation(type="list", formula1='"Correlation,Covariance"',
                            showDropDown=False)
dv_yesno  = DataValidation(type="list", formula1='"Yes,No"', showDropDown=False)
ws.add_data_validation(dv_matrix)
ws.add_data_validation(dv_yesno)
dv_matrix.sqref = "B7"
dv_yesno.sqref  = "B9"

thin_border(ws, 4, 1, 11, 3)

# ── Section B: Factor Labels ──────────────────────────────────────────────────
ws.row_dimensions[13].height = 20
ws.merge_cells("A13:F13")
section_hdr(ws["A13"], "B.  FACTOR LABELS  (auto-populated from DATA Row 4 — edit Display Name)")

ws.row_dimensions[14].height = 18
for col, hdr_text in enumerate(["Factor #", "DATA Header (auto)", "Display Name", "Color Tag"], start=1):
    sub_hdr(ws.cell(row=14, column=col), hdr_text)

for i in range(1, 11):
    r = 14 + i
    ws.row_dimensions[r].height = 16
    label(ws.cell(row=r, column=1), f"Factor {i}", bold=True)
    # Auto-pull header from DATA row 4 (col C = col 3, so offset by 2)
    c = ws.cell(row=r, column=2)
    c.value = f'=IFERROR(DATA!{get_column_letter(i+2)}$4,"")'
    c.font = Font(name="Calibri", size=10, color="595959")
    c.alignment = Alignment(horizontal="center")
    # Display name — user editable
    c2 = ws.cell(row=r, column=3)
    c2.value = f"=B{r}"   # defaults to DATA header; user can override
    c2.font = Font(name="Calibri", bold=True, size=10, color=C_NAVY)
    c2.fill = PatternFill("solid", fgColor="FFF2CC")
    c2.alignment = Alignment(horizontal="center")
    # Color tag (for charts)
    c3 = ws.cell(row=r, column=4)
    c3.value = ""
    c3.fill = PatternFill("solid", fgColor=C_LGRAY)

thin_border(ws, 14, 1, 24, 4)

# ── Section C: PC Labels ──────────────────────────────────────────────────────
ws.row_dimensions[26].height = 20
ws.merge_cells("A26:F26")
section_hdr(ws["A26"], "C.  PRINCIPAL COMPONENT LABELS  (rename your PCs here)")

ws.row_dimensions[27].height = 18
for col, hdr_text in enumerate(
        ["PC #", "Default Name", "Custom Label", "Economic Interpretation", "Include in PCA Reg?"],
        start=1):
    sub_hdr(ws.cell(row=27, column=col), hdr_text)

for i in range(1, 11):
    r = 27 + i
    ws.row_dimensions[r].height = 16
    label(ws.cell(row=r, column=1), f"PC{i}", bold=True)
    c2 = ws.cell(row=r, column=2)
    c2.value = f"PC{i}"
    c2.font = Font(name="Calibri", size=10, color="595959")
    c2.alignment = Alignment(horizontal="center")
    # Custom label — user editable
    c3 = ws.cell(row=r, column=3)
    c3.value = f"PC{i}"
    c3.font = Font(name="Calibri", bold=True, size=10, color=C_NAVY)
    c3.fill = PatternFill("solid", fgColor="FFF2CC")
    c3.alignment = Alignment(horizontal="center")
    # Economic interpretation
    c4 = ws.cell(row=r, column=4)
    c4.value = ""
    c4.fill = PatternFill("solid", fgColor=C_LGRAY)
    # Include in PCA Reg
    c5 = ws.cell(row=r, column=5)
    c5.value = "Yes" if i <= 5 else "No"
    c5.font = Font(name="Calibri", bold=True, size=10, color=C_NAVY)
    c5.fill = PatternFill("solid", fgColor="FFF2CC")
    c5.alignment = Alignment(horizontal="center")

dv_yesno2 = DataValidation(type="list", formula1='"Yes,No"', showDropDown=False)
ws.add_data_validation(dv_yesno2)
dv_yesno2.sqref = "E28:E37"
thin_border(ws, 27, 1, 37, 5)

# Color legend
ws.row_dimensions[39].height = 16
ws.merge_cells("A39:B39")
ws["A39"].value = "Yellow cells = user-editable"
ws["A39"].font = Font(name="Calibri", italic=True, size=9)
c_swatch = ws.cell(row=39, column=3)
c_swatch.fill = PatternFill("solid", fgColor="FFF2CC")

freeze(ws, "A4")


# ══════════════════════════════════════════════════════════════════════════════
# 3. PCA OUTPUT SHEET
# ══════════════════════════════════════════════════════════════════════════════
ws = sheets["PCA"]
ws.sheet_view.showGridLines = False

ws.row_dimensions[1].height = 30
ws.merge_cells("A1:N1")
hdr_style(ws["A1"], "PRINCIPAL COMPONENT ANALYSIS OUTPUT", bg=C_NAVY, size=14)

# Run button instruction row
ws.row_dimensions[2].height = 22
ws.merge_cells("A2:N2")
ws["A2"].value = (
    "Click RUN PCA (button below) or press Ctrl+Shift+P to execute. "
    "All tables are populated by VBA — do not edit directly."
)
ws["A2"].font = Font(name="Calibri", italic=True, color=C_BLUE, size=9)
ws["A2"].fill = PatternFill("solid", fgColor="EBF3FB")
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

set_col_width(ws, {
    "A": 22, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14,
    "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14,
    "M": 14, "N": 14,
})

# ── Section 1: Summary Statistics ────────────────────────────────────────────
ws.row_dimensions[4].height = 20
ws.merge_cells("A4:F4")
section_hdr(ws["A4"], "1.  SUMMARY STATISTICS")

ws.row_dimensions[5].height = 18
for col, hdr_text in enumerate(["Factor", "N", "Mean", "Std Dev", "Min", "Max"], start=1):
    sub_hdr(ws.cell(row=5, column=col), hdr_text)

for i in range(1, 11):
    r = 5 + i
    ws.row_dimensions[r].height = 15
    c = ws.cell(row=r, column=1)
    c.value = f'=IFERROR(CONFIG!C{14+i},"")'   # Display name from CONFIG
    c.font = Font(name="Calibri", bold=True, size=10)
    for col in range(2, 7):
        ws.cell(row=r, column=col).number_format = "0.0000"
        if (r - 5) % 2 == 0:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=C_LGRAY)

thin_border(ws, 5, 1, 15, 6)

# ── Section 2: Correlation / Covariance Matrix ─────────────────────────────
ws.row_dimensions[17].height = 20
ws.merge_cells("A17:K17")
section_hdr(ws["A17"], "2.  CORRELATION / COVARIANCE MATRIX  (set in CONFIG)")

ws.row_dimensions[18].height = 18
sub_hdr(ws.cell(row=18, column=1), "")
for i in range(1, 11):
    sub_hdr(ws.cell(row=18, column=i + 1), f"F{i}")

for i in range(1, 11):
    r = 18 + i
    ws.row_dimensions[r].height = 15
    sub_hdr(ws.cell(row=r, column=1), f"F{i}")
    for j in range(1, 11):
        c = ws.cell(row=r, column=j + 1)
        c.number_format = "0.0000"
        if i == j:
            c.fill = PatternFill("solid", fgColor="E2EFDA")
        elif (i + j) % 2 == 0:
            c.fill = PatternFill("solid", fgColor=C_LGRAY)

thin_border(ws, 18, 1, 28, 11)

# ── Section 3: Eigenvalue Table ───────────────────────────────────────────────
ws.row_dimensions[30].height = 20
ws.merge_cells("A30:F30")
section_hdr(ws["A30"], "3.  EIGENVALUES  &  VARIANCE EXPLAINED")

ws.row_dimensions[31].height = 18
for col, hdr_text in enumerate(
        ["PC", "Eigenvalue", "% Variance", "Cumulative %", "Status", "Label (CONFIG)"],
        start=1):
    sub_hdr(ws.cell(row=31, column=col), hdr_text)

for i in range(1, 11):
    r = 31 + i
    ws.row_dimensions[r].height = 15
    c = ws.cell(row=r, column=1)
    c.value = f'=IFERROR(CONFIG!C{27+i},"")'   # PC label from CONFIG
    c.font = Font(name="Calibri", bold=True, size=10)
    ws.cell(row=r, column=2).number_format = "0.0000"
    ws.cell(row=r, column=3).number_format = "0.00%"
    ws.cell(row=r, column=4).number_format = "0.00%"
    ws.cell(row=r, column=5).value = ""   # VBA writes "Retained" / "Dropped"
    c6 = ws.cell(row=r, column=6)
    c6.value = f'=IFERROR(CONFIG!D{27+i},"")'
    c6.font = Font(name="Calibri", italic=True, size=9, color=C_BLUE)
    if i % 2 == 0:
        for col in range(2, 7):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=C_LGRAY)

thin_border(ws, 31, 1, 41, 6)

# ── Section 4: Factor Loadings ────────────────────────────────────────────────
ws.row_dimensions[43].height = 20
ws.merge_cells("A43:N43")
section_hdr(ws["A43"], "4.  FACTOR LOADINGS  (eigenvectors × √eigenvalue)")

ws.row_dimensions[44].height = 18
sub_hdr(ws.cell(row=44, column=1), "Factor \\ PC")
for i in range(1, 11):
    sub_hdr(ws.cell(row=44, column=i + 1), f"=IFERROR(CONFIG!C{27+i},\"PC{i}\")")

for i in range(1, 11):
    r = 44 + i
    ws.row_dimensions[r].height = 15
    c = ws.cell(row=r, column=1)
    c.value = f'=IFERROR(CONFIG!C{14+i},"")'
    c.font = Font(name="Calibri", bold=True, size=10)
    for j in range(1, 11):
        c2 = ws.cell(row=r, column=j + 1)
        c2.number_format = "0.0000"
        if (i + j) % 2 == 0:
            c2.fill = PatternFill("solid", fgColor=C_LGRAY)

thin_border(ws, 44, 1, 54, 11)

# ── Section 5: PC Scores header (VBA fills in data) ───────────────────────────
ws.row_dimensions[56].height = 20
ws.merge_cells("A56:N56")
section_hdr(ws["A56"], "5.  PC SCORES  (time series)  — populated by RUN PCA")

ws.row_dimensions[57].height = 18
sub_hdr(ws.cell(row=57, column=1), "Date")
for i in range(1, 11):
    sub_hdr(ws.cell(row=57, column=i + 1), f"=IFERROR(CONFIG!C{27+i},\"PC{i}\")")

freeze(ws, "A5")


# ══════════════════════════════════════════════════════════════════════════════
# 4. REGRESSION SHEET
# ══════════════════════════════════════════════════════════════════════════════
ws = sheets["REGRESSION"]
ws.sheet_view.showGridLines = False

ws.row_dimensions[1].height = 30
ws.merge_cells("A1:H1")
hdr_style(ws["A1"], "OLS REGRESSION  —  Dependent Variable on Original Factors", bg=C_NAVY, size=14)

ws.row_dimensions[2].height = 22
ws.merge_cells("A2:H2")
ws["A2"].value = "Click RUN REGRESSION or press Ctrl+Shift+R. All cells populated by VBA."
ws["A2"].font = Font(name="Calibri", italic=True, color=C_BLUE, size=9)
ws["A2"].fill = PatternFill("solid", fgColor="EBF3FB")
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

set_col_width(ws, {"A": 24, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14})

# ── Model Summary ─────────────────────────────────────────────────────────────
ws.row_dimensions[4].height = 20
ws.merge_cells("A4:H4")
section_hdr(ws["A4"], "1.  MODEL SUMMARY")

summary_labels = [
    (5,  "Dependent Variable",  ""),
    (6,  "N (Observations)",    ""),
    (7,  "R-Squared",           ""),
    (8,  "Adjusted R-Squared",  ""),
    (9,  "F-Statistic",         ""),
    (10, "Prob (F-Stat)",       ""),
    (11, "AIC",                 ""),
    (12, "BIC",                 ""),
]
for row, lbl_text, _ in summary_labels:
    ws.row_dimensions[row].height = 16
    sub_hdr(ws.cell(row=row, column=1), lbl_text)
    c = ws.cell(row=row, column=2)
    c.number_format = "0.0000"
    c.font = Font(name="Calibri", bold=True, size=10, color=C_NAVY)
    c.alignment = Alignment(horizontal="center")

thin_border(ws, 5, 1, 12, 2)

# ── ANOVA Table ───────────────────────────────────────────────────────────────
ws.row_dimensions[14].height = 20
ws.merge_cells("A14:H14")
section_hdr(ws["A14"], "2.  ANOVA TABLE")

ws.row_dimensions[15].height = 18
for col, hdr_text in enumerate(["Source", "SS", "df", "MS", "F", "Prob(F)"], start=1):
    sub_hdr(ws.cell(row=15, column=col), hdr_text)

for i, src in enumerate(["Regression", "Residual", "Total"], start=1):
    r = 15 + i
    ws.row_dimensions[r].height = 15
    label(ws.cell(row=r, column=1), src, bold=(i == 3))
    for col in range(2, 7):
        ws.cell(row=r, column=col).number_format = "0.0000"
        if i % 2 == 0:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=C_LGRAY)

thin_border(ws, 15, 1, 18, 6)

# ── Coefficients Table ────────────────────────────────────────────────────────
ws.row_dimensions[20].height = 20
ws.merge_cells("A20:H20")
section_hdr(ws["A20"], "3.  COEFFICIENTS")

ws.row_dimensions[21].height = 18
for col, hdr_text in enumerate(
        ["Variable", "Coefficient", "Std Error", "t-Stat", "p-Value",
         "CI Lower", "CI Upper", "Significance"],
        start=1):
    sub_hdr(ws.cell(row=21, column=col), hdr_text)

for i in range(0, 12):  # intercept + up to 11 factors
    r = 22 + i
    ws.row_dimensions[r].height = 15
    c = ws.cell(row=r, column=1)
    c.value = "Intercept" if i == 0 else f'=IFERROR(CONFIG!C{14+i},"")'
    c.font = Font(name="Calibri", bold=(i == 0), size=10)
    for col in range(2, 9):
        ws.cell(row=r, column=col).number_format = "0.0000"
        if i % 2 == 0:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=C_LGRAY)

thin_border(ws, 21, 1, 33, 8)

# ── Residuals Section ─────────────────────────────────────────────────────────
ws.row_dimensions[35].height = 20
ws.merge_cells("A35:H35")
section_hdr(ws["A35"], "4.  RESIDUALS  (populated by VBA)")

ws.row_dimensions[36].height = 18
for col, hdr_text in enumerate(["Date", "Actual", "Fitted", "Residual", "Std Residual"], start=1):
    sub_hdr(ws.cell(row=36, column=col), hdr_text)

freeze(ws, "A4")


# ══════════════════════════════════════════════════════════════════════════════
# 5. PCA_REG SHEET
# ══════════════════════════════════════════════════════════════════════════════
ws = sheets["PCA_REG"]
ws.sheet_view.showGridLines = False

ws.row_dimensions[1].height = 30
ws.merge_cells("A1:H1")
hdr_style(ws["A1"], "PCA REGRESSION  —  Dependent Variable on Principal Components", bg=C_NAVY, size=14)

ws.row_dimensions[2].height = 22
ws.merge_cells("A2:H2")
ws["A2"].value = (
    "Click RUN PCA REGRESSION or Ctrl+Shift+Q. "
    "PCs used are controlled by CONFIG Section C column E."
)
ws["A2"].font = Font(name="Calibri", italic=True, color=C_BLUE, size=9)
ws["A2"].fill = PatternFill("solid", fgColor="EBF3FB")
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

set_col_width(ws, {"A": 24, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14})

# Model Summary
ws.row_dimensions[4].height = 20
ws.merge_cells("A4:H4")
section_hdr(ws["A4"], "1.  PC REGRESSION MODEL SUMMARY")
for i, lbl_text in enumerate(
        ["Dependent Variable", "PCs Included", "N", "R-Squared",
         "Adjusted R-Squared", "F-Statistic", "Prob (F-Stat)"],
        start=5):
    ws.row_dimensions[i].height = 16
    sub_hdr(ws.cell(row=i, column=1), lbl_text)
    c = ws.cell(row=i, column=2)
    c.font = Font(name="Calibri", bold=True, size=10, color=C_NAVY)
    c.alignment = Alignment(horizontal="center")
thin_border(ws, 5, 1, 11, 2)

# PC Coefficients
ws.row_dimensions[13].height = 20
ws.merge_cells("A13:H13")
section_hdr(ws["A13"], "2.  PC COEFFICIENTS")
ws.row_dimensions[14].height = 18
for col, hdr_text in enumerate(
        ["PC", "Label", "Coeff", "Std Error", "t-Stat", "p-Value", "CI Lower", "CI Upper"],
        start=1):
    sub_hdr(ws.cell(row=14, column=col), hdr_text)
for i in range(0, 11):
    r = 15 + i
    ws.row_dimensions[r].height = 15
    c = ws.cell(row=r, column=1)
    c.value = "Intercept" if i == 0 else f"PC{i}"
    c.font = Font(name="Calibri", bold=(i == 0), size=10)
    if i > 0:
        c2 = ws.cell(row=r, column=2)
        c2.value = f'=IFERROR(CONFIG!C{27+i},"")'
        c2.font = Font(name="Calibri", italic=True, size=10, color=C_BLUE)
    for col in range(3, 9):
        ws.cell(row=r, column=col).number_format = "0.0000"
        if i % 2 == 0:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=C_LGRAY)
thin_border(ws, 14, 1, 25, 8)

# Back-transformation table
ws.row_dimensions[27].height = 20
ws.merge_cells("A27:H27")
section_hdr(ws["A27"], "3.  BACK-TRANSFORMED FACTOR CONTRIBUTIONS  (via PC loadings)")
ws.row_dimensions[28].height = 18
for col, hdr_text in enumerate(["Factor", "Display Name", "PC Contribution", "% of R²"], start=1):
    sub_hdr(ws.cell(row=28, column=col), hdr_text)
for i in range(1, 11):
    r = 28 + i
    ws.row_dimensions[r].height = 15
    label(ws.cell(row=r, column=1), f"Factor {i}")
    c = ws.cell(row=r, column=2)
    c.value = f'=IFERROR(CONFIG!C{14+i},"")'
    for col in range(3, 5):
        ws.cell(row=r, column=col).number_format = "0.0000"
        if i % 2 == 0:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=C_LGRAY)
thin_border(ws, 28, 1, 38, 4)

freeze(ws, "A4")


# ══════════════════════════════════════════════════════════════════════════════
# 6. CHARTS SHEET
# ══════════════════════════════════════════════════════════════════════════════
ws = sheets["CHARTS"]
ws.sheet_view.showGridLines = False

ws.row_dimensions[1].height = 30
ws.merge_cells("A1:P1")
hdr_style(ws["A1"], "CHARTS  —  Populated by RUN PCA / RUN REGRESSION", bg=C_NAVY, size=14)

ws.row_dimensions[2].height = 18
ws.merge_cells("A2:P2")
ws["A2"].value = (
    "Charts are created / refreshed automatically when you run the analysis. "
    "Scree plot | Cumulative Variance | Factor Loadings Heatmap | PC Scores | Residuals"
)
ws["A2"].font = Font(name="Calibri", italic=True, color=C_BLUE, size=9)
ws["A2"].fill = PatternFill("solid", fgColor="EBF3FB")
ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

# Placeholder labels for chart positions
chart_placeholders = [
    (4,  "A", "P", "SCREE PLOT"),
    (26, "A", "H", "CUMULATIVE VARIANCE EXPLAINED"),
    (26, "I", "P", "PC SCORES (Time Series)"),
    (48, "A", "P", "FACTOR LOADINGS HEATMAP"),
]
for row, c1, c2, title in chart_placeholders:
    ws.row_dimensions[row].height = 20
    ws.merge_cells(f"{c1}{row}:{c2}{row}")
    section_hdr(ws[f"{c1}{row}"], title)
    # Placeholder area
    for r in range(row + 1, row + 20):
        ws.row_dimensions[r].height = 14
        if r == row + 1:
            cell_addr = f"{c1}{r}"
            col_end = c2
            ws.merge_cells(f"{cell_addr}:{col_end}{r}")
            ws[cell_addr].value = "(chart auto-populated by VBA)"
            ws[cell_addr].font = Font(name="Calibri", italic=True, size=10, color=C_MGRAY)
            ws[cell_addr].alignment = Alignment(horizontal="center", vertical="center")


# ══════════════════════════════════════════════════════════════════════════════
# NAMED RANGES (workbook-level)
# ══════════════════════════════════════════════════════════════════════════════
from openpyxl.workbook.defined_name import DefinedName

def add_named_range(wb, name, formula):
    dn = DefinedName(name, attr_text=formula)
    wb.defined_names[name] = dn

# VBA uses these by name
add_named_range(wb, "DATA_START_ROW", "CONFIG!$B$10")
add_named_range(wb, "DATA_END_ROW",   "CONFIG!$B$11")
add_named_range(wb, "NUM_PCS",        "CONFIG!$B$6")
add_named_range(wb, "MATRIX_TYPE",    "CONFIG!$B$7")
add_named_range(wb, "CONF_LEVEL",     "CONFIG!$B$8")
add_named_range(wb, "INCLUDE_INT",    "CONFIG!$B$9")


# ══════════════════════════════════════════════════════════════════════════════
# WRITE OUTPUT
# ══════════════════════════════════════════════════════════════════════════════
out_path = "PCA_Analytics.xlsx"
wb.save(out_path)
print(f"Workbook saved: {out_path}")
print()
print("Next steps:")
print("  1. Rename to PCA_Analytics.xlsm in Excel (File → Save As → .xlsm)")
print("  2. Alt+F11 → File → Import File → import each file in vba/")
print("  3. Import order: modUtils.bas, modPCA.bas, modRegression.bas, modCharts.bas, modMain.bas")
print("  4. Add buttons: Developer → Insert → Button → assign macro RunAll")
print()
print("VBA keyboard shortcuts (set in modMain):")
print("  Ctrl+Shift+P  → RunPCA")
print("  Ctrl+Shift+R  → RunRegression")
print("  Ctrl+Shift+Q  → RunPCAReg")
print("  Ctrl+Shift+A  → RunAll")
