Attribute VB_Name = "MasterDashboardPopup"
' =====================================================================================
' MasterDashboardPopup
'
' Feature: "Add Feature (box display & pop-up box both option) in Symbol column"
'
' The Python app already gives you the "box display" half of this for free (a small
' tooltip with a handful of key fields, appears the instant you click a Symbol cell,
' no macros needed).
'
' This module gives you the OTHER half: a full pop-up card (every single column,
' not just a condensed set) the moment you click any Symbol cell in Master_Dashboard-8.
'
' Unlike a hard-coded version (which breaks the moment columns get reordered/added/
' removed via the app's column sequencer), this one reads the header row itself at
' click-time, so it always matches whatever columns/order are actually on the sheet.
'
' SETUP: paste this into a new module (Alt+F11 > Insert > Module) — NOT into
' "ThisWorkbook" or the sheet's own code module — then follow the wiring step below.
' Requires the Macro-Enabled Workbook format (.xlsm), same as MasterDashboardBuilder.bas.
' =====================================================================================

Public Sub ShowMasterDashboardCard(ByVal ws As Worksheet, ByVal r As Long)
    Dim lastCol As Long, lastRow As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If r < 2 Or r > lastRow Then Exit Sub
    If Trim(ws.Cells(r, 1).Value) = "" Then Exit Sub

    Dim msg As String
    msg = "==========================================================" & vbCrLf & _
          "                     STOCK DETAIL CARD                     " & vbCrLf & _
          "==========================================================" & vbCrLf & vbCrLf

    For c = 1 To lastCol
        Dim label As String, val As String
        label = Trim(ws.Cells(1, c).Value)
        If label <> "" Then
            val = ws.Cells(r, c).Text
            msg = msg & label & ": " & val & vbCrLf
        End If
    Next c

    MsgBox msg, vbInformation, "Data View: " & ws.Cells(r, 1).Value
End Sub
