# Master_Dashboard-8 — running it inside Excel itself

This gives Excel the ability to rebuild `Master_Dashboard-8` on its own,
straight from whatever is currently in the source tabs — useful if someone
hand-edits data after the Python script has already produced the workbook.

## One-time setup

1. Open the workbook produced by the Python script in Excel.
2. **File > Save As** → choose **Excel Macro-Enabled Workbook (*.xlsm)**.
   (Plain `.xlsx` cannot store macros — this is the one unavoidable tradeoff.)
3. Press **Alt+F11** to open the VBA editor.
4. **File > Import File...** and select `MasterDashboardBuilder.bas`.
   This adds a module containing the full build logic.
5. In the **Project** pane (left side), double-click **ThisWorkbook**, and
   paste this in:

   ```vb
   Private Sub Workbook_Open()
       BuildMasterDashboard
   End Sub
   ```

6. Save (Ctrl+S), keeping the Macro-Enabled format.
7. Close and reopen the file. If prompted, click **Enable Content** /
   **Enable Macros**. `Master_Dashboard-8` will rebuild automatically,
   right after **Main Tab**, every time the workbook opens.

## Running it manually any time

You don't have to close/reopen to trigger a rebuild — press **Alt+F8**,
select `BuildMasterDashboard`, and click **Run**.

## What it does

Same logic as the Python version: scans the first 15 rows of each source
tab for a header row, fuzzy-matches column names against the alias list,
joins everything onto `Symbol` (first non-blank value wins per column),
and writes the result into a freshly rebuilt `Master_Dashboard-8` sheet
with the light-magenta header, Arial font, and the same number formats.

## Note on the two "auto-run" paths

- **Python script** (`app_R15_with_master_dashboard.py`): builds
  `Master_Dashboard-8` automatically the moment you click
  "🚀 Execute Structural Consolidation" — this is the source of truth,
  since it runs against the freshly-loaded, freshly-cleaned data.
- **This VBA macro**: a *secondary* safety net for if the file is opened
  and edited directly in Excel afterward. It rebuilds from whatever is
  currently sitting in the tabs at that moment — so if you've since
  deleted or altered rows in Excel, the dashboard will reflect *that*,
  not the original Python output. Keep that in mind if the two ever look
  different from each other.
