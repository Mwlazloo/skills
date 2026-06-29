# Excel PCA & Regression Analysis Tool

A full-featured GUI for Principal Component Analysis and OLS Regression built entirely in Excel VBA — no Python, no add-ins, no installation required beyond Excel itself.

## What It Does

| Feature | Details |
|---|---|
| **PCA** | Jacobi eigendecomposition on covariance or correlation matrix |
| **Rename PCs** | Give each component a meaningful economic/domain name |
| **Regression** | OLS on any combination of PC scores or raw ranges |
| **Session history** | Every PCA and regression is saved; recall, review, or delete anytime |
| **Export** | Loadings, scores, and regression tables go to their own worksheets |
| **GUI** | Five pop-up forms — no macros to remember |

---

## Setup (One-Time, ~2 Minutes)

### Step 1 — Download the source files

Download the `src/` folder from this repo. You need all six `.bas` files.

### Step 2 — Enable VBA project trust *(required once)*

Excel must be allowed to manipulate its own VBA project so the installer can create the form windows.

1. `File → Options → Trust Center → Trust Center Settings…`
2. Click **Macro Settings**
3. Check **"Trust access to the VBA project object model"**
4. Click OK → OK

> You can uncheck this again after setup — it is only needed during installation.

### Step 3 — Open a new macro-enabled workbook

`File → New → Blank Workbook`, then save it as **`PCA_Tool.xlsm`** (macro-enabled format).

### Step 4 — Import all six modules

Press **Alt + F11** to open the VBA Editor, then for each `.bas` file:

`File → Import File…` → select the file → Open

Import in this order (or any order — they all work):

| File | Purpose |
|---|---|
| `modMatrix.bas` | Matrix algebra (multiply, inverse, Jacobi eigen) |
| `modPCA.bas` | PCA computation and data access |
| `modRegression.bas` | OLS computation |
| `modStorage.bas` | Hidden-sheet persistence layer |
| `modMain.bas` | Entry points and utilities |
| `modSetup.bas` | One-time installer that builds all GUI forms |

### Step 5 — Run the installer

Back in the VBA Editor, press **F5** (or go to `Run → Run Macro`) and run:

```
SetupPCATool
```

The installer will:
- Create a **Data** worksheet for your input data
- Build five GUI forms automatically
- Open the main form immediately

> After setup, you can re-disable "Trust access to the VBA project object model" in Trust Center.

### Step 6 — Launch anytime

Press **Alt + F8** → select `ShowMainForm` → Run.

Or add a button to any sheet: `Developer tab → Insert → Button` and assign `ShowMainForm`.

---

## How to Use

### Running a PCA

1. Paste your data into the **Data** sheet (variables in columns, observations in rows).
2. From the main form, click **"Run New PCA Analysis…"**
3. Click **"..."** to select your data range.
4. Choose options:
   - **Standardize** (on by default) uses the correlation matrix — good when variables have different units.
   - **Max components** — leave blank to keep all.
5. Give the session a name, click **Run PCA**.

The results form opens with three tabs:
- **Variance Explained** — eigenvalues, % variance, cumulative %. Select a PC and rename it.
- **Loadings** — variable loadings on each component. Copy to sheet for further analysis.
- **Scores (Preview)** — first 20 rows of PC scores. "Export All Scores" sends everything to a sheet.

### Renaming PCs

On the Variance Explained tab, click a component in the list, type a name in the box, click **Rename PC**. Names persist across sessions.

### Running a Regression

From the main form: **"Run New Regression…"**, or from the PCA results form: **"Run Regression with These PCs"**.

**Using PC scores** (most common):
1. Select the PCA session from the dropdown.
2. Move desired PCs from "Available" → "Selected" with the **Add >>** button.
3. Pick your Y range (dependent variable).
4. Click **Run Regression**.

**Using a custom range** (any data):
- Switch to "Use Custom Range" and select your X matrix range directly.

### Results

The regression results form shows:
- R², Adjusted R², F-statistic and p-value
- Full coefficient table: estimate, standard error, t-statistic, p-value, significance stars
- **Export Results to Sheet** copies everything to a new worksheet named `Reg_REG_XXX`

### Recalling Previous Work

The main form lists all saved PCA and regression sessions. Click any session → **Open** to re-view. Sessions survive workbook close/reopen. Click **Delete** to remove a session permanently.

---

## Technical Notes

### Algorithms

**PCA:** Jacobi iterative eigendecomposition of the sample covariance (or correlation) matrix. Converges for all real symmetric matrices. Maximum iterations = 200 × n².

**OLS Regression:** Normal equations via Gauss-Jordan matrix inversion with partial pivoting. Standard errors from `MSE × diag[(X'X)⁻¹]`. t-statistics and p-values via Excel's `TDIST`. F-statistic via `FDIST`.

### Storage

Each analysis session is stored as a very-hidden worksheet (`Visible = xlSheetVeryHidden`). These sheets are invisible to the user but persist with the workbook. An index sheet `_INDEX` tracks all sessions.

Session IDs follow the format `PCA_001`, `REG_002`, etc.

### Limitations

- **No multicollinearity check** — if your X matrix is (near-)singular, the inverse will fail with a message.
- **No robust standard errors** — uses conventional OLS SEs. Heteroskedasticity-consistent SEs are not implemented.
- **Jacobi convergence** — rare failure on pathological matrices (near-zero off-diagonal elements that oscillate). In practice this does not occur with real financial or economic data.
- **Memory** — very large datasets (>5,000 obs, >50 variables) may be slow due to VBA's lack of BLAS. Tested comfortably up to ~2,000 × 30.

---

## File Reference

```
src/
├── modMatrix.bas        Matrix operations (multiply, inverse, Jacobi eigen, standardize)
├── modPCA.bas           PCA: run, store, retrieve, rename, export
├── modRegression.bas    OLS: run from PCA scores or custom range, coefficients & stats
├── modStorage.bas       Hidden-sheet persistence, session index
├── modMain.bas          ShowMainForm(), Auto_Open(), utilities
└── modSetup.bas         One-time installer: builds all 5 UserForms via VBProject API
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `SetupPCATool` says "Cannot access VBA project" | Enable Trust Center setting (Step 2) |
| Form doesn't open after setup | Run `ShowMainForm` from Alt+F8 |
| "Singular matrix" error in regression | PCs are uncorrelated by construction — this indicates a custom X range with collinear columns |
| Session list is empty after reopening | Re-run `InitStorage` once from Alt+F8 (only needed if workbook was restructured) |
| Scores don't match expected sign | PC signs are arbitrary; the explained variance is correct regardless of sign |
