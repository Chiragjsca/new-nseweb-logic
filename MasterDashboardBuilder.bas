Attribute VB_Name = "MasterDashboardBuilder"
Option Explicit

' =====================================================================================
' Master_Dashboard-8 builder — native VBA port of the Python/Apps Script logic.
' Joins every listed tab on Symbol into one wide reference sheet.
'
' SETUP (one-time):
'   1. Open the workbook in Excel, then File > Save As > "Excel Macro-Enabled
'      Workbook (*.xlsm)" — plain .xlsx cannot store macros.
'   2. Press Alt+F11 to open the VBA editor.
'   3. File > Import File... and select this .bas file (adds this module).
'   4. In the Project pane, double-click "ThisWorkbook" and paste the small
'      Workbook_Open snippet from the accompanying instructions (or run
'      BuildMasterDashboard manually any time via Alt+F8).
'   5. Save. On every future open, Excel will rebuild Master_Dashboard-8 from
'      whatever is currently in the source tabs.
' =====================================================================================

Private Const MASTER_SHEET_NAME As String = "Master_Dashboard-8"
Private Const HEADER_SCAN_ROWS As Long = 15
Private Const HEADER_FILL_RGB As Long = 15121130   ' RGB(234,209,220) = EAD1DC

' ----------------------------- FIELD MAP -----------------------------
' Each row: Array(Label, SheetName, "Alias1|Alias2|...", Format, IsKey)
Private Function GetFieldMap() As Variant
    GetFieldMap = Array( _
        Array("Symbol", "BhavCopy_NSE_CM", "TckrSymb|SYMBOL|Symb", "text", True), _
        Array("ISIN", "BhavCopy_NSE_CM", "ISIN|ISIN NUMBER", "text", False), _
        Array("Series", "BhavCopy_NSE_CM", "SctySrs|SERIES|Series|Srs", "text", False), _
        Array("Company Name (Capital)", "BhavCopy_NSE_CM", "FinInstrmNm|NAME OF COMPANY|Name Of Company|Security Name|SECURITY|Security|COMPANY NAME|COMPANY'S NAME|Company Name|Company's Name", "text", False), _
        Array("Company Name", "EQUITY_L|SME_EQUITY_L", "NAME OF COMPANY|NAME_OF_COMPANY|COMPANY'S_NAME|Name Of Company|Security Name|SECURITY|Security|COMPANY NAME|COMPANY'S NAME|Company Name|Company's Name", "text", False), _
        Array("Date of Listing", "EQUITY_L|SME_EQUITY_L", "DATE OF LISTING|DATE_OF_LISTING|LISTING_DATE|LISTING DATE", "date", False), _
        Array("Trade Date", "BhavCopy_NSE_CM", "TradDt|Trade Date", "date", False), _
        Array("Segment", "BhavCopy_NSE_CM", "Src", "text", False), _
        Array("Market Lot", "BhavCopy_NSE_CM", "NewBrdLotQty|MARKET LOT|Market Lot", "qty", False), _
        Array("T0 Tag", "Eligible_T0_Securities", "SERIES|SctySrs|Srs|Series", "text", False), _
        Array("Remarks", "sec_list", "Remarks", "text", False), _
        Array("Face Value", "EQUITY_L|SME_EQUITY_L", "FACE VALUE|Face Value(Rs.)|FACE_VALUE|Face Value (Rs.)", "price", False), _
        Array("No. of Trades", "BhavCopy_NSE_CM", "TtlNbOfTxsExctd|No. of Trades|NO OF TRADES|TRADES|Trade|NO_OF_TRADES", "qty", False), _
        Array("Traded Qty", "BhavCopy_NSE_CM", "TtlTradgVol|TTL TRD QNTY|TRADED QUANTITY|NET_TRDQTY|Traded Qty|NET TRD QTY|NET TRDQTY|TTL_TRD_QNTY", "qty", False), _
        Array("Delivery Qty", "sec_bhavdata_full", "DELIV QTY|DELIV QUANTITY|Delivery quantity|DELIVERY QNTY|DELIV_QNTY|DELIV QNTY|DELIV_QTY", "qty", False), _
        Array("Turnover (Rs.)", "BhavCopy_NSE_CM", "TtlTrfVal|NET_TRDVAL|NET_TRD_VAL|NET TRD VAL|NET TRDVAL|Turnover (Rs.)|NET TRADED VALUE|Net Traded Value|Traded Value", "qty", False), _
        Array("Issue Size", "mcap", "Issue Size", "qty", False), _
        Array("Mkt Cap (Rs. Crores)", "StocksTraded", "Mkt Cap (Rs Crores)|Mkt Cap (₹ Crores)|Market Cap (₹ Crores)", "crores", False), _
        Array("Market Cap(Rs.)", "mcap", "Market Cap(Rs.)|Mkt Cap(Rs.)|Market Cap (Rs.)|Mkt Cap (Rs.)|Market Cap(Rs)|Mkt Cap(Rs)|Market Cap (Rs)|Mkt Cap (Rs)", "qty", False), _
        Array("Delivery %", "sec_bhavdata_full", "DELIV PER|DELIV %|delivery percentage|Delivery Percentage (%)|DELIV_PER", "percent", False), _
        Array("Value (Rs. Crores)", "StocksTraded", "Value (Rs Crores)|Value (₹ Crores)", "crores", False), _
        Array("Value (Rs.)", "BhavCopy_NSE_CM", "TtlTrfVal|NET_TRDVAL", "qty", False), _
        Array("Volume (Lakhs)", "StocksTraded", "Volume (Lakhs)", "lakhs", False), _
        Array("Volume", "BhavCopy_NSE_CM", "TtlTradgVol|NET_TRDQTY", "qty", False), _
        Array("Band", "sec_list", "Band", "number", False), _
        Array("% Change", "StocksTraded", "%chng|% Change", "percent", False), _
        Array("Close Price", "BhavCopy_NSE_CM", "ClsPric|CLOSE PRICE|Close Price|CLOSE_PRICE", "price", False), _
        Array("CMP/LTP", "BhavCopy_NSE_CM", "LastPric|LAST PRICE|Last Price|LTP|LAST_PRICE", "price", False), _
        Array("Prev Close", "BhavCopy_NSE_CM", "PrvsClsgPric|PREV CLOSE|Previous close|PREV_CL_PR|PREV_CLOSE", "price", False), _
        Array("Open (Rs.)", "BhavCopy_NSE_CM", "OpnPric|Open Price|OPEN PRICE|OPEN_PRICE", "price", False), _
        Array("High (Rs.)", "BhavCopy_NSE_CM", "HghPric|HIGH PRICE|High Price|HIGH_PRICE", "price", False), _
        Array("Low (Rs.)", "BhavCopy_NSE_CM", "LwPric|Low Price|LOW PRICE|LOW_PRICE", "price", False), _
        Array("52W High", "CM_52_wk_High_low", "Adjusted_52_Week_High|52_Week_High|52W_High|52 Week High|52W High|HI_52_WK", "price", False), _
        Array("52W High Date", "CM_52_wk_High_low", "52_Week_High_Date|52 Week High Date|52_Week_High_DT|52W High Date|52 W High Date|52 W High Dt.|52W High Dt.", "date", False), _
        Array("52W Low", "CM_52_wk_High_low", "Adjusted_52_Week_Low|52 Week Low|52_Week_Low|52W_Low|52W Low|LO_52_WK", "price", False), _
        Array("52W Low Date", "CM_52_wk_High_low", "52_Week_Low_DT|52 Week Low Date|52 W Low Date|52 W Low Dt.|52W Low Dt.", "date", False), _
        Array("Symbol P/E", "PE", "SYMBOL P/E|Symbol P/E", "ratio", False), _
        Array("Adjusted P/E", "PE", "ADJUSTED P/E|Adjusted P/E", "ratio", False), _
        Array("T0 Effective Date", "Eligible_T0_Securities", "Effective Date", "text", False), _
        Array("Paid Up Value", "EQUITY_L|SME_EQUITY_L", "PAID UP VALUE|PAID_UP_VALUE", "price", False), _
        Array("Category", "mcap", "Category", "text", False) _
    )
End Function

Private Function GetSymbolAliases() As Variant
    GetSymbolAliases = Split("SYMBOL|TckrSymb|Symb|Symbol", "|")
End Function

Private Function GetNumberFormat(fmt As String) As String
    Select Case fmt
        Case "price": GetNumberFormat = "#,##0.00"
        Case "qty": GetNumberFormat = "#,##0"
        Case "date": GetNumberFormat = "dd-mmm-yyyy"
        Case "percent": GetNumberFormat = "0.00" & Chr(34) & "%" & Chr(34)
        Case "ratio": GetNumberFormat = "0.00"
        Case "crores": GetNumberFormat = "#,##0.00" & Chr(34) & " Cr" & Chr(34)
        Case "lakhs": GetNumberFormat = "#,##0.00" & Chr(34) & " L" & Chr(34)
        Case "number": GetNumberFormat = "0"
        Case Else: GetNumberFormat = "@"
    End Select
End Function

' ----------------------------- HELPERS -----------------------------

Private Function NormalizeHeader(v As Variant) As String
    Dim s As String
    If IsError(v) Then
        NormalizeHeader = ""
        Exit Function
    End If
    If IsEmpty(v) Or IsNull(v) Then
        NormalizeHeader = ""
        Exit Function
    End If
    s = LCase(Trim(CStr(v)))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeHeader = s
End Function

Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not s Is Nothing
End Function

Private Function FindHeaderRow(ws As Worksheet, aliasesDict As Object) As Long
    Dim ur As Range: Set ur = ws.UsedRange
    Dim startRow As Long: startRow = ur.Row
    Dim startCol As Long: startCol = ur.Column
    Dim endCol As Long: endCol = startCol + ur.Columns.Count - 1
    Dim maxScanRow As Long
    maxScanRow = startRow + WorksheetFunction.Min(HEADER_SCAN_ROWS, ur.Rows.Count) - 1

    Dim bestRow As Long, bestScore As Long, r As Long, c As Long, score As Long, norm As String
    bestRow = -1: bestScore = 0
    For r = startRow To maxScanRow
        score = 0
        For c = startCol To endCol
            norm = NormalizeHeader(ws.Cells(r, c).Value)
            If norm <> "" Then
                If aliasesDict.Exists(norm) Then score = score + 1
            End If
        Next c
        If score > bestScore Then
            bestScore = score
            bestRow = r
        End If
    Next r
    FindHeaderRow = IIf(bestScore > 0, bestRow, -1)
End Function

Private Function BuildHeaderIndex(ws As Worksheet, headerRow As Long) As Object
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    Dim ur As Range: Set ur = ws.UsedRange
    Dim startCol As Long: startCol = ur.Column
    Dim endCol As Long: endCol = startCol + ur.Columns.Count - 1
    Dim c As Long, norm As String
    For c = startCol To endCol
        norm = NormalizeHeader(ws.Cells(headerRow, c).Value)
        If norm <> "" Then
            If Not idx.Exists(norm) Then idx.Add norm, c
        End If
    Next c
    Set BuildHeaderIndex = idx
End Function

Private Function MatchColumn(headerIndex As Object, aliasesPipe As String) As Long
    Dim aliasArr As Variant: aliasArr = Split(aliasesPipe, "|")
    Dim a As Variant, norm As String
    For Each a In aliasArr
        norm = NormalizeHeader(a)
        If headerIndex.Exists(norm) Then
            MatchColumn = headerIndex(norm)
            Exit Function
        End If
    Next a
    MatchColumn = -1
End Function

Private Sub SortStringArray(arr() As String)
    Dim i As Long, j As Long, temp As String
    For i = LBound(arr) + 1 To UBound(arr)
        temp = arr(i)
        j = i - 1
        Do While j >= LBound(arr) And arr(j) > temp
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = temp
    Next i
End Sub

' ----------------------------- MAIN BUILD -----------------------------

Public Sub BuildMasterDashboard()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim fieldMap As Variant: fieldMap = GetFieldMap()
    Dim symbolAliases As Variant: symbolAliases = GetSymbolAliases()
    Dim lo As Long, hi As Long
    lo = LBound(fieldMap): hi = UBound(fieldMap)

    ' Every alias used anywhere, normalized, for header-row detection.
    Dim allAliases As Object: Set allAliases = CreateObject("Scripting.Dictionary")
    Dim a As Variant, i As Long
    For Each a In symbolAliases
        allAliases(NormalizeHeader(a)) = True
    Next a
    For i = lo To hi
        Dim aliasArr As Variant: aliasArr = Split(fieldMap(i)(2), "|")
        For Each a In aliasArr
            allAliases(NormalizeHeader(a)) = True
        Next a
    Next i

    ' Group field indices by source sheet name.
    Dim sheetFieldIdx As Object: Set sheetFieldIdx = CreateObject("Scripting.Dictionary")
    For i = lo To hi
        Dim sn As String: sn = fieldMap(i)(1)
        If Not sheetFieldIdx.Exists(sn) Then
            Dim newColl As Collection: Set newColl = New Collection
            sheetFieldIdx.Add sn, newColl
        End If
        sheetFieldIdx(sn).Add i
    Next i

    Dim masterData As Object: Set masterData = CreateObject("Scripting.Dictionary")  ' Symbol -> Dictionary(label->value)
    Dim symbolOrder As Collection: Set symbolOrder = New Collection

    Dim sKey As Variant
    For Each sKey In sheetFieldIdx.Keys
        Dim shName As String: shName = CStr(sKey)
        If SheetExists(wb, shName) Then
            Dim ws As Worksheet: Set ws = wb.Sheets(shName)
            Dim headerRow As Long: headerRow = FindHeaderRow(ws, allAliases)
            If headerRow > 0 Then
                Dim hIdx As Object: Set hIdx = BuildHeaderIndex(ws, headerRow)
                Dim symCol As Long: symCol = MatchColumn(hIdx, Join(symbolAliases, "|"))
                If symCol > 0 Then
                    Dim ur As Range: Set ur = ws.UsedRange
                    Dim lastDataRow As Long: lastDataRow = ur.Row + ur.Rows.Count - 1
                    Dim fieldIdxColl As Collection: Set fieldIdxColl = sheetFieldIdx(shName)
                    Dim r As Long
                    For r = headerRow + 1 To lastDataRow
                        Dim symRaw As Variant: symRaw = ws.Cells(r, symCol).Value
                        If Not IsEmpty(symRaw) Then
                            If Trim(CStr(symRaw)) <> "" Then
                                Dim sym As String: sym = Trim(CStr(symRaw))
                                Dim rec As Object
                                If masterData.Exists(sym) Then
                                    Set rec = masterData(sym)
                                Else
                                    Set rec = CreateObject("Scripting.Dictionary")
                                    masterData.Add sym, rec
                                    symbolOrder.Add sym
                                End If

                                Dim fIdxV As Variant
                                For Each fIdxV In fieldIdxColl
                                    Dim fi As Long: fi = CLng(fIdxV)
                                    Dim lbl As String: lbl = fieldMap(fi)(0)
                                    Dim isKeyField As Boolean: isKeyField = CBool(fieldMap(fi)(4))
                                    If isKeyField Then
                                        rec(lbl) = sym
                                    Else
                                        Dim colIdx As Long: colIdx = MatchColumn(hIdx, fieldMap(fi)(2))
                                        If colIdx > 0 Then
                                            Dim curV As Variant
                                            If rec.Exists(lbl) Then
                                                curV = rec(lbl)
                                            Else
                                                curV = Empty
                                            End If
                                            If IsEmpty(curV) Then
                                                rec(lbl) = ws.Cells(r, colIdx).Value
                                            ElseIf VarType(curV) = vbString And curV = "" Then
                                                rec(lbl) = ws.Cells(r, colIdx).Value
                                            End If
                                        End If
                                    End If
                                Next fIdxV
                            End If
                        End If
                    Next r
                End If
            End If
        End If
    Next sKey

    WriteMasterSheet wb, fieldMap, masterData, symbolOrder

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Master_Dashboard-8 build error: " & Err.Description, vbExclamation
    End If
End Sub

Private Sub WriteMasterSheet(wb As Workbook, fieldMap As Variant, masterData As Object, symbolOrder As Collection)
    Dim lo As Long, hi As Long
    lo = LBound(fieldMap): hi = UBound(fieldMap)
    Dim nFields As Long: nFields = hi - lo + 1

    Application.DisplayAlerts = False
    If SheetExists(wb, MASTER_SHEET_NAME) Then wb.Sheets(MASTER_SHEET_NAME).Delete
    Application.DisplayAlerts = True

    Dim ws As Worksheet: Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = MASTER_SHEET_NAME

    Dim c As Long
    For c = 1 To nFields
        ws.Cells(1, c).Value = fieldMap(c - 1 + lo)(0)
    Next c
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nFields))
        .Interior.Color = HEADER_FILL_RGB
        .Font.Bold = True
        .Font.Name = "Arial"
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(204, 204, 204)
    End With

    Dim n As Long: n = symbolOrder.Count
    Dim symArr() As String
    ReDim symArr(1 To n)
    Dim k As Long
    For k = 1 To n
        symArr(k) = symbolOrder(k)
    Next k
    If n > 0 Then SortStringArray symArr

    Dim rIdx As Long: rIdx = 2
    For k = 1 To n
        Dim sym2 As String: sym2 = symArr(k)
        Dim rec2 As Object: Set rec2 = masterData(sym2)
        For c = 1 To nFields
            Dim lbl2 As String: lbl2 = fieldMap(c - 1 + lo)(0)
            Dim fmt2 As String: fmt2 = fieldMap(c - 1 + lo)(3)
            Dim v2 As Variant
            If rec2.Exists(lbl2) Then
                v2 = rec2(lbl2)
            Else
                v2 = ""
            End If
            With ws.Cells(rIdx, c)
                .Value = v2
                .NumberFormat = GetNumberFormat(fmt2)
                .Font.Name = "Arial"
            End With
        Next c
        rIdx = rIdx + 1
    Next k

    If rIdx > 2 Then
        With ws.Range(ws.Cells(1, 1), ws.Cells(rIdx - 1, nFields))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(204, 204, 204)
        End With
    End If

    ws.Columns.AutoFit
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ' Move it right after the first sheet (e.g. "Main Tab"), matching the Python build.
    If wb.Sheets.Count > 1 Then
        ws.Move After:=wb.Sheets(1)
    End If
End Sub
