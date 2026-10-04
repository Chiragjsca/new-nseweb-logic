# Optional: full pop-up card for Master_Dashboard-8 (bonus, macro-based)

The Python app (`app_updated.py`) already gives you a no-macro **"box display"**:
click any Symbol cell in `Master_Dashboard-8` and a small tooltip with the key
fields (CMP/LTP, % Change, Close Price, Prev Close, Volume, Mkt Cap, P/E) pops
up instantly. Nothing to install — it's baked into the generated `.xlsx`.

If you also want the **full pop-up card** (every column, not just the
condensed set) shown in the sample from your Word doc, that does require a
macro — Excel has no way to run code on cell-click without one. This is
provided as an optional add-on, exactly the way `MasterDashboardBuilder.bas`
already works in `EXCEL_MACRO_SETUP.md`.

## Setup

1. Open the generated workbook, save as **Excel Macro-Enabled Workbook (.xlsm)**.
2. Alt+F11 → **File > Import File...** → select `MasterDashboardPopup.bas`.
   (This adds a *standard* module — leave it as-is.)
3. In the **Project** pane, double-click the `Master_Dashboard-8` sheet itself
   (not "ThisWorkbook", not a standard module — the sheet's own code page),
   and paste:

   ```vb
   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
       If Target.Cells.Count = 1 And Target.Column = 1 Then
           ShowMasterDashboardCard Me, Target.Row
       End If
   End Sub
   ```

4. Save (keep .xlsm format), close and reopen, click **Enable Content**.
5. Click any Symbol cell in column A — the full detail card pops up.

Unlike a hard-coded version, `ShowMasterDashboardCard` reads the header row
itself at click time, so it keeps working even if you reorder, add, or remove
Master_Dashboard-8 columns using the app's column sequencer.
