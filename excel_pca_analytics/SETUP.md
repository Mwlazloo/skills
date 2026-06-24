# PCA Analytics Workbook — Setup Guide

## What's Included

| File | Purpose |
|---|---|
| `build_workbook.py` | Generates `PCA_Analytics.xlsx` (sheet structure, formatting, named ranges) |
| `vba/modUtils.bas` | Matrix ops: multiply, transpose, Gauss-Jordan inverse, standardize, sort |
| `vba/modPCA.bas` | PCA via Jacobi eigendecomposition + output writer |
| `vba/modRegression.bas` | OLS regression + PCA regression + back-transformation |
| `vba/modCharts.bas` | Scree plot, cumulative variance, loadings bar chart, PC scores, residuals |
| `vba/modMain.bas` | Entry points, keyboard shortcuts, reset, diagnostics, button installer |

---

## First-Time Setup (5 minutes)

### Step 1 — Build the workbook structure
```
pip install openpyxl
python build_workbook.py
```
This creates `PCA_Analytics.xlsx`.

### Step 2 — Enable macros
1. Open `PCA_Analytics.xlsx` in Excel
2. **File → Save As → Excel Macro-Enabled Workbook (.xlsm)**  
   Name it `PCA_Analytics.xlsm`

### Step 3 — Import VBA modules
1. **Alt+F11** to open VBA Editor
2. **File → Import File** — import in this order:
   - `vba/modUtils.bas`
   - `vba/modPCA.bas`
   - `vba/modRegression.bas`
   - `vba/modCharts.bas`
   - `vba/modMain.bas`
3. Close VBA Editor, **Ctrl+S** to save

### Step 4 — Add buttons (optional but recommended)
1. **Alt+F11** → Immediate Window (Ctrl+G)
2. Type: `modMain.AddSheetButtons` and press Enter
3. Buttons appear on every sheet automatically

---

## Workbook Architecture

```
DATA        ← All raw data lives here (CIQ PRO formulas go here)
CONFIG      ← Settings + factor/PC renaming (yellow cells = editable)
PCA         ← Eigenvalues, loadings, PC scores (VBA-populated)
REGRESSION  ← OLS output: coefficients, t-stats, p-values, residuals
PCA_REG     ← PCA regression + back-transformation to factors
CHARTS      ← Scree plot, cumulative variance, loadings, PC scores, residuals
```

---

## DATA Sheet

Column layout (row 4 = headers, row 5+ = data):

| Col A | Col B | Col C | Col D | … |
|---|---|---|---|---|
| Date | Dep Var Return | Factor 1 | Factor 2 | … |

**Replace placeholder cells with CIQ PRO formulas**, e.g.:
```
=CIQ("SP500","IQ_TOTAL_RETURN","IQ_FY-1")
```
For a time series, use `CIQRANGE` and paste-transpose into the column.

**To add a factor:** insert a column (C–L), add a header in row 4 — CONFIG auto-detects it.  
**To remove a factor:** delete the column — all analyses adapt on the next run.

---

## CONFIG Sheet

### Section A — Analysis Settings
| Setting | Default | Notes |
|---|---|---|
| Dependent Variable Name | Portfolio Return | Label only — shown in outputs |
| PCs to Retain | 5 | Controls scree chart + PCA_REG |
| Matrix Type | Correlation | `Correlation` = standardized; `Covariance` = raw |
| Confidence Level | 0.95 | Used for regression CI bands |
| Include Intercept | Yes | OLS intercept toggle |

### Section B — Factor Labels
- **DATA Header (auto)**: pulled from DATA row 4 — do not edit
- **Display Name**: rename factors for cleaner output (e.g. "Mkt-RF" → "Market Premium")

### Section C — PC Labels
- **Custom Label**: rename PCs after inspection (e.g. "PC1" → "Macro Risk")
- **Economic Interpretation**: free-text notes
- **Include in PCA Reg?**: `Yes/No` — controls which PCs enter the regression

---

## Running Analyses

### Keyboard Shortcuts
| Shortcut | Action |
|---|---|
| `Ctrl+Shift+A` | Run everything (PCA + both regressions) |
| `Ctrl+Shift+P` | Run PCA only |
| `Ctrl+Shift+R` | Run OLS regression only |
| `Ctrl+Shift+Q` | Run PCA regression only |
| `Ctrl+Shift+X` | Reset all outputs |

### Typical Workflow
1. Paste / refresh data on **DATA** sheet
2. Set factor display names in **CONFIG Section B**
3. `Ctrl+Shift+P` → inspect scree plot, note eigenvalues > 1
4. Rename retained PCs in **CONFIG Section C**
5. Mark which PCs to include in regression (column E)
6. `Ctrl+Shift+Q` → check PCA_REG back-transformation
7. `Ctrl+Shift+R` → compare vs OLS on raw factors

---

## Charts Produced

| Chart | Sheet | What to look for |
|---|---|---|
| **Scree Plot** | CHARTS | Elbow in eigenvalues; 80% cumulative threshold line |
| **Cumulative Variance** | CHARTS | How many PCs needed for X% explained |
| **Factor Loadings** | CHARTS | Which factors drive each PC (bold = \|loading\| ≥ 0.5) |
| **PC Scores (Time Series)** | CHARTS | PC dynamics over time |
| **Residuals vs Fitted** | CHARTS | Homoscedasticity check |

---

## Swapping Factor Sets

1. On **DATA**: overwrite headers in row 4 and replace formulas in the factor columns
2. Run `modMain.RefreshFactorNames` (or use the Refresh button on CONFIG) — this re-syncs display names for unchanged factors
3. Re-run analysis (`Ctrl+Shift+A`)

---

## Troubleshooting

**"X'X is singular"** — multicollinearity. Remove a correlated factor or use PCA Regression.  
**"Not enough data rows"** — need ≥ 3 observations; check CONFIG rows 10–11.  
**No factor columns found** — ensure DATA row 4 has headers starting at column C.  
**Charts show wrong PC count** — update CONFIG Section A "PCs to Retain" before re-running.
