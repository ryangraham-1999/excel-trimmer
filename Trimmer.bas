Attribute VB_Name = "Trimmer"
' The trimmer (designs/stage-B-trimmer.md 4.2 to 4.9).
' TrimStock is the "Trim stock" button: it finds the export among the open workbooks by its
' headings, asks which one, and calls TrimSheet. TrimSheet is the trim itself: it reads the
' export sheet (never changing it) and writes the result into a new tab "Trimmed" in the
' export's workbook. The self-test (module TrimmerSelfTest) calls the same TrimSheet.
' Nothing is saved, and Trimmer.xlsm's own sheet is never changed.
Option Explicit

' Shared with the self-test (4.9).
Public Const KEPT_NAMES As String = "Veh Description|Reg Number|Mileage|Selling Price|Disposal Method"
Public Const RESULT_TAB As String = "Trimmed"
Public Const NONE_TEXT As String = "none"

' The block names, in order (4.5). Other is always last and is not in the list.
Private Const BLOCKS As String = "Retail|Demo|Courtesy|Auction"
Private Const OTHER_BLOCK As String = "Other"

Private Const KEPT_COUNT As Long = 5
Private Const RESULT_COLS As Long = 7
Private Const MAX_DESCRIPTION_WIDTH As Double = 60
Private Const TICK_BOX_WIDTH As Double = 4
Private Const MSG_TITLE As String = "Trimmer"

' ---------------------------------------------------------------------------------------------
' The button (4.2)
' ---------------------------------------------------------------------------------------------

Public Sub TrimStock()
    Dim candidates As Collection, ws As Worksheet, result As Worksheet
    Dim i As Long, listText As String, answer As String, message As String

    On Error GoTo Failed
    Set candidates = FindCandidates()
    If candidates.Count = 0 Then
        MsgBox "No open workbook has the stock headings in row 1. Open the Keyloop export, then press Trim stock again.", vbExclamation, MSG_TITLE
        Exit Sub
    ElseIf candidates.Count = 1 Then
        Set ws = candidates(1)
        If MsgBox("Trim """ & ws.Name & """ in """ & ws.Parent.Name & """?", vbYesNo + vbQuestion, MSG_TITLE) <> vbYes Then Exit Sub
    Else
        For i = 1 To candidates.Count
            listText = listText & vbLf & i & "  " & candidates(i).Name & " in " & candidates(i).Parent.Name
        Next i
        answer = InputBox("More than one open sheet has the stock headings. Type the number of the one to trim:" & vbLf & listText, MSG_TITLE, "1")
        i = ListedNumber(answer, candidates.Count)
        If i = 0 Then Exit Sub
        Set ws = candidates(i)
    End If

    Set result = TrimSheet(ws, False, message)
    If result Is Nothing Then Exit Sub

    ' 4.2 step 5: show the new tab, headings frozen at the top (4.6). Showing is not part of the trim.
    On Error Resume Next
    result.Parent.Activate
    result.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    MsgBox message, vbInformation, MSG_TITLE
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "The trimmer stopped: " & Err.Description, vbExclamation, MSG_TITLE
End Sub

' Every worksheet of every open, visible workbook (not this file, not PERSONAL.XLSB, not a
' Trimmed tab) whose row 1 has all five kept headings (4.2 step 1).
Private Function FindCandidates() As Collection
    Dim found As New Collection, wb As Workbook, ws As Worksheet
    For Each wb In Application.Workbooks
        If Not (wb Is ThisWorkbook) And UCase$(wb.Name) <> "PERSONAL.XLSB" Then
            If IsVisible(wb) Then
                For Each ws In wb.Worksheets
                    If Not IsResultTabName(ws.Name) Then
                        If HasAllKeptHeadings(ws) Then found.Add ws
                    End If
                Next ws
            End If
        End If
    Next wb
    Set FindCandidates = found
End Function

Private Function IsVisible(ByVal wb As Workbook) As Boolean
    Dim w As Window
    For Each w In wb.Windows
        If w.Visible Then
            IsVisible = True
            Exit Function
        End If
    Next w
End Function

' "Trimmed" or "Trimmed (<n>)".
Private Function IsResultTabName(ByVal tabName As String) As Boolean
    Dim inside As String
    If StrComp(tabName, RESULT_TAB, vbTextCompare) = 0 Then
        IsResultTabName = True
    ElseIf StrComp(Left$(tabName, Len(RESULT_TAB) + 2), RESULT_TAB & " (", vbTextCompare) = 0 And Right$(tabName, 1) = ")" Then
        inside = Mid$(tabName, Len(RESULT_TAB) + 3, Len(tabName) - Len(RESULT_TAB) - 3)
        IsResultTabName = Len(inside) > 0 And Not (inside Like "*[!0-9]*")
    End If
End Function

' The number typed into the picker, if it is one of 1 to count; else 0.
Private Function ListedNumber(ByVal answer As String, ByVal count As Long) As Long
    answer = Trim$(answer)
    If Len(answer) = 0 Or Len(answer) > 6 Or answer Like "*[!0-9]*" Then Exit Function
    If CLng(answer) >= 1 And CLng(answer) <= count Then ListedNumber = CLng(answer)
End Function

' ---------------------------------------------------------------------------------------------
' The trim (4.2 steps 2 to 4, 4.3 to 4.8)
' ---------------------------------------------------------------------------------------------

' Trims the export sheet ws into a new tab of its workbook and returns that tab. When the trim
' stops instead, nothing is made, it returns Nothing and message holds why. On success message
' holds the final message (4.2 step 5). quiet:=True never shows a message box (the self-test).
Public Function TrimSheet(ByVal ws As Worksheet, Optional ByVal quiet As Boolean = False, Optional ByRef message As String = "") As Worksheet
    Dim screenWas As Boolean, wb As Workbook, newWs As Worksheet, problem As String
    Dim lastCol As Long, lastRow As Long, keptCol(1 To KEPT_COUNT) As Long
    Dim values() As Variant, formats() As String, rowCount As Long
    Dim blockOfRow() As Long, order() As Long, blockCount(1 To 5) As Long, printOk As Boolean
    Dim desc As String

    message = ""
    screenWas = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Set wb = ws.Parent

    ' 4.2 step 2, 4.3: the headings.
    problem = HeadingProblems(ws, keptCol, lastCol)
    If problem <> "" Then GoTo Stopped

    ' 4.2 step 3, 4.4: the vehicle rows, into memory.
    lastRow = LastUsedRow(ws, lastCol)
    rowCount = ReadVehicleRows(ws, keptCol, lastCol, lastRow, values, formats)
    If rowCount = 0 Then
        problem = "The export has no vehicles under its headings."
        GoTo Stopped
    End If

    ' 4.5: blocks and order.
    OrderRows values, rowCount, blockOfRow, order, blockCount

    ' 4.2 step 4: the new tab, found by its name.
    Set newWs = AddResultTab(wb)
    If newWs Is Nothing Then
        problem = "The trimmer stopped: could not add the Trimmed tab."
        GoTo Stopped
    End If

    ' 4.6 to 4.8.
    WriteResult newWs, values, formats, rowCount, blockOfRow, order, blockCount
    printOk = SetUpPrinting(newWs)

    message = FinalMessage(newWs.Name, rowCount, blockCount)
    If Not printOk Then message = message & " (print set-up skipped)"
    Application.ScreenUpdating = screenWas
    Set TrimSheet = newWs
    Exit Function

Stopped:
    Application.ScreenUpdating = screenWas
    message = problem
    If Not quiet Then MsgBox message, vbExclamation, MSG_TITLE
    Set TrimSheet = Nothing
    Exit Function

Failed:
    desc = Err.Description
    Resume CleanUp
CleanUp:
    On Error Resume Next
    If Not newWs Is Nothing Then
        Application.DisplayAlerts = False
        newWs.Delete
        Application.DisplayAlerts = True
    End If
    Application.ScreenUpdating = True
    On Error GoTo 0
    message = "The trimmer stopped: " & desc
    If Not quiet Then MsgBox message, vbExclamation, MSG_TITLE
    Set TrimSheet = Nothing
End Function

' 4.3: "" when every kept name matches exactly one heading of row 1, else the message listing
' every problem. Also gives each kept name's column and the last column of row 1.
Private Function HeadingProblems(ByVal ws As Worksheet, ByRef keptCol() As Long, ByRef lastCol As Long) As String
    Dim missing As String, twice As String
    HeadingCheck ws, keptCol, lastCol, missing, twice
    HeadingProblems = Trim$(missing & " " & twice)
End Function

' 4.2 step 1: a candidate's row 1 has all five kept headings (a doubled one is found at step 2).
Private Function HasAllKeptHeadings(ByVal ws As Worksheet) As Boolean
    Dim keptCol(1 To KEPT_COUNT) As Long, lastCol As Long, missing As String, twice As String
    HeadingCheck ws, keptCol, lastCol, missing, twice
    HasAllKeptHeadings = (missing = "")
End Function

' 4.3: matches row 1 against the kept names. missing is "The export is missing: ..." or "";
' twice is one "The export has ... twice (columns ...)." per doubled name, or "".
Private Sub HeadingCheck(ByVal ws As Worksheet, ByRef keptCol() As Long, ByRef lastCol As Long, ByRef missing As String, ByRef twice As String)
    Dim names() As String, k As Long, c As Long, n As Long, cols As String, letters() As String

    names = Split(KEPT_NAMES, "|")
    lastCol = LastHeadingColumn(ws)
    For k = 1 To KEPT_COUNT
        n = 0
        cols = ""
        keptCol(k) = 0
        For c = 1 To lastCol
            If StrComp(Trim$(CellText(ws.Cells(1, c).Value)), names(k - 1), vbTextCompare) = 0 Then
                n = n + 1
                If n = 1 Then keptCol(k) = c
                cols = cols & "|" & ColumnLetter(c)
            End If
        Next c
        If n = 0 Then
            missing = missing & IIf(missing = "", "", ", ") & names(k - 1)
        ElseIf n >= 2 Then
            letters = Split(Mid$(cols, 2), "|")
            twice = twice & IIf(twice = "", "", " ") & "The export has """ & names(k - 1) & """ " & _
                    IIf(n = 2, "twice", n & " times") & " (columns " & AndList(letters) & ")."
        End If
    Next k
    If missing <> "" Then missing = "The export is missing: " & missing & "."
End Sub

' "H", "H and K", "H, J and K".
Private Function AndList(items() As String) As String
    Dim i As Long, last As Long
    last = UBound(items)
    For i = LBound(items) To last
        If i = LBound(items) Then
            AndList = items(i)
        ElseIf i = last Then
            AndList = AndList & " and " & items(i)
        Else
            AndList = AndList & ", " & items(i)
        End If
    Next i
End Function

Private Function ColumnLetter(ByVal col As Long) As String
    Dim n As Long
    n = col
    Do While n > 0
        ColumnLetter = Chr$(65 + (n - 1) Mod 26) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

' 4.3: the last non-empty cell of row 1.
Private Function LastHeadingColumn(ByVal ws As Worksheet) As Long
    Dim maxCol As Long
    maxCol = ws.Columns.Count
    If Not IsEmpty(ws.Cells(1, maxCol).Value) Then
        LastHeadingColumn = maxCol
    Else
        LastHeadingColumn = ws.Cells(1, maxCol).End(xlToLeft).Column
    End If
End Function

' 4.3: the larger of UsedRange's last row and each column's End(xlUp), up to lastCol.
Private Function LastUsedRow(ByVal ws As Worksheet, ByVal lastCol As Long) As Long
    Dim c As Long, r As Long, maxRow As Long
    maxRow = ws.Rows.Count
    LastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For c = 1 To lastCol
        If Not IsEmpty(ws.Cells(maxRow, c).Value) Then
            r = maxRow
        Else
            r = ws.Cells(maxRow, c).End(xlUp).Row
        End If
        If r > LastUsedRow Then LastUsedRow = r
    Next c
End Function

' 4.4: the five kept cells (value and number format) of every vehicle row, in export order.
' values(i, k) and formats(i, k), i = 1 to the count returned, k = 1 to 5 in kept-column order.
Private Function ReadVehicleRows(ByVal ws As Worksheet, keptCol() As Long, ByVal lastCol As Long, ByVal lastRow As Long, _
                                 ByRef values() As Variant, ByRef formats() As String) As Long
    Dim data As Variant, r As Long, c As Long, n As Long, k As Long, blank As Boolean, total As Long

    If lastRow < 2 Then Exit Function
    total = lastRow - 1
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
    ReDim values(1 To total, 1 To KEPT_COUNT)
    ReDim formats(1 To total, 1 To KEPT_COUNT)
    For r = 1 To total
        blank = True
        For c = 1 To lastCol
            If Not IsBlankValue(data(r, c)) Then
                blank = False
                Exit For
            End If
        Next c
        If Not blank Then
            n = n + 1
            For k = 1 To KEPT_COUNT
                values(n, k) = data(r, keptCol(k))
                formats(n, k) = ws.Cells(r + 1, keptCol(k)).NumberFormat
            Next k
        End If
    Next r
    ReadVehicleRows = n
End Function

' A blank cell: empty, or text of only spaces (2).
Private Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Len(Replace(v, " ", "")) = 0)
    End If
End Function

' 4.5: each row's block, and the order of the rows (indexes into values) block by block:
' rows with a reg by reg key (stable), then rows with no reg in export order.
Private Sub OrderRows(values() As Variant, ByVal rowCount As Long, ByRef blockOfRow() As Long, ByRef order() As Long, ByRef blockCount() As Long)
    Dim keys() As String, i As Long, b As Long, n As Long, pos As Long, first As Long
    Dim withReg() As Long, tmp() As Long, nReg As Long

    ReDim keys(1 To rowCount)
    ReDim blockOfRow(1 To rowCount)
    ReDim order(1 To rowCount)
    ReDim withReg(1 To rowCount)
    ReDim tmp(1 To rowCount)
    For i = 1 To rowCount
        keys(i) = RegKey(values(i, 2))
        blockOfRow(i) = BlockOf(values(i, 5))
        blockCount(blockOfRow(i)) = blockCount(blockOfRow(i)) + 1
    Next i

    pos = 0
    For b = 1 To 5
        first = pos + 1
        nReg = 0
        For i = 1 To rowCount
            If blockOfRow(i) = b And Len(keys(i)) > 0 Then
                nReg = nReg + 1
                withReg(nReg) = i
            End If
        Next i
        If nReg > 1 Then MergeSort withReg, keys, 1, nReg, tmp
        For n = 1 To nReg
            pos = pos + 1
            order(pos) = withReg(n)
        Next n
        For i = 1 To rowCount
            If blockOfRow(i) = b And Len(keys(i)) = 0 Then
                pos = pos + 1
                order(pos) = i
            End If
        Next i
    Next b
End Sub

' Stable merge sort of idx(lo..hi) by keys(idx(.)), character code order (4.5).
Private Sub MergeSort(idx() As Long, keys() As String, ByVal lo As Long, ByVal hi As Long, tmp() As Long)
    Dim m As Long, i As Long, j As Long, k As Long
    If hi <= lo Then Exit Sub
    m = (lo + hi) \ 2
    MergeSort idx, keys, lo, m, tmp
    MergeSort idx, keys, m + 1, hi, tmp
    i = lo
    j = m + 1
    k = lo
    Do While i <= m And j <= hi
        If StrComp(keys(idx(j)), keys(idx(i)), vbBinaryCompare) < 0 Then
            tmp(k) = idx(j)
            j = j + 1
        Else
            tmp(k) = idx(i)
            i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= m
        tmp(k) = idx(i)
        i = i + 1
        k = k + 1
    Loop
    Do While j <= hi
        tmp(k) = idx(j)
        j = j + 1
        k = k + 1
    Loop
    For k = lo To hi
        idx(k) = tmp(k)
    Next k
End Sub

' 4.5: the reg key: the reg as text, upper case, every space, tab and character 160 removed.
Private Function RegKey(ByVal v As Variant) As String
    RegKey = UCase$(CellText(v))
    RegKey = Replace(RegKey, " ", "")
    RegKey = Replace(RegKey, vbTab, "")
    RegKey = Replace(RegKey, Chr$(160), "")
End Function

' 4.5: 1 to 4 for the names in BLOCKS (trimmed, any case), 5 for Other.
Private Function BlockOf(ByVal v As Variant) As Long
    Dim names() As String, i As Long, t As String
    names = Split(BLOCKS, "|")
    t = Trim$(CellText(v))
    For i = 0 To UBound(names)
        If StrComp(t, names(i), vbTextCompare) = 0 Then
            BlockOf = i + 1
            Exit Function
        End If
    Next i
    BlockOf = UBound(names) + 2
End Function

Private Function BlockName(ByVal b As Long) As String
    Dim names() As String
    names = Split(BLOCKS, "|")
    If b <= UBound(names) + 1 Then BlockName = names(b - 1) Else BlockName = OTHER_BLOCK
End Function

' A cell value as text: a number in its shortest form with "." as the point (whatever the
' PC's settings), anything else as VBA writes it; empty and error cells give "".
Private Function CellText(ByVal v As Variant) As String
    If IsEmpty(v) Or IsError(v) Or IsNull(v) Then
        CellText = ""
    ElseIf VarType(v) = vbString Then
        CellText = v
    ElseIf IsNumeric(v) And VarType(v) <> vbBoolean Then
        CellText = NumberText(CDbl(v))
    ElseIf VarType(v) = vbDate Then
        CellText = NumberText(CDbl(v))
    Else
        CellText = CStr(v)
    End If
End Function

' The shortest decimal form of a number: "." as the point, no thousands separator, no exponent
' (5301, 30354, 0, 13495.01, 0.5).
Private Function NumberText(ByVal d As Double) As String
    Dim s As String, sign As String, mant As String, ex As Long, p As Long, digits As String, pointAt As Long
    s = Trim$(Str$(d))
    If Left$(s, 1) = "-" Then
        sign = "-"
        s = Mid$(s, 2)
    End If
    p = InStr(1, s, "E", vbTextCompare)
    If p > 0 Then
        mant = Left$(s, p - 1)
        ex = CLng(Mid$(s, p + 1))
    Else
        mant = s
    End If
    p = InStr(mant, ".")
    If p > 0 Then
        digits = Left$(mant, p - 1) & Mid$(mant, p + 1)
        pointAt = p - 1
    Else
        digits = mant
        pointAt = Len(mant)
    End If
    pointAt = pointAt + ex
    If pointAt <= 0 Then
        digits = String$(1 - pointAt, "0") & digits
        pointAt = 1
    ElseIf pointAt > Len(digits) Then
        digits = digits & String$(pointAt - Len(digits), "0")
    End If
    s = Left$(digits, pointAt)
    mant = Mid$(digits, pointAt + 1)
    Do While Len(mant) > 0 And Right$(mant, 1) = "0"
        mant = Left$(mant, Len(mant) - 1)
    Loop
    Do While Len(s) > 1 And Left$(s, 1) = "0"
        s = Mid$(s, 2)
    Loop
    If Len(s) = 0 Then s = "0"
    If Len(mant) > 0 Then s = s & "." & mant
    If s = "0" Then sign = ""
    NumberText = sign & s
End Function

' 4.2 step 4 (amended): add a sheet after the last one and find it by its name not being among
' the names before; never by what Worksheets.Add returns, its position or the active sheet.
Private Function AddResultTab(ByVal wb As Workbook) As Worksheet
    Dim before As String, sh As Worksheet, found As Worksheet, n As Long, countBefore As Long

    before = "|"
    For Each sh In wb.Worksheets
        before = before & sh.Name & "|"
    Next sh
    countBefore = wb.Worksheets.Count
    wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    For Each sh In wb.Worksheets
        If InStr(1, before, "|" & sh.Name & "|", vbBinaryCompare) = 0 Then
            Set found = sh
            n = n + 1
        End If
    Next sh
    If n <> 1 Or wb.Worksheets.Count <> countBefore + 1 Then Exit Function
    found.Name = FreeTabName(wb)
    Set AddResultTab = found
End Function

' "Trimmed", or "Trimmed (2)", "Trimmed (3)" ... the first name no sheet of wb has.
Private Function FreeTabName(ByVal wb As Workbook) As String
    Dim n As Long, candidate As String
    candidate = RESULT_TAB
    n = 1
    Do While SheetNameTaken(wb, candidate)
        n = n + 1
        candidate = RESULT_TAB & " (" & n & ")"
    Loop
    FreeTabName = candidate
End Function

Private Function SheetNameTaken(ByVal wb As Workbook, ByVal tabName As String) As Boolean
    Dim sh As Object
    For Each sh In wb.Sheets
        If StrComp(sh.Name, tabName, vbTextCompare) = 0 Then
            SheetNameTaken = True
            Exit Function
        End If
    Next sh
End Function

' 4.6 and 4.7: the heading row, then the blocks.
Private Sub WriteResult(ByVal rs As Worksheet, values() As Variant, formats() As String, ByVal rowCount As Long, _
                        blockOfRow() As Long, order() As Long, blockCount() As Long)
    Dim names() As String, r As Long, b As Long, lastBlock As Long, pos As Long, i As Long, firstRow As Long
    Dim resultCol As Variant, k As Long

    names = Split(KEPT_NAMES, "|")
    resultCol = Array(1, 2, 4, 5, 7)

    ' Row 1: the headings, once (C7); C and F blank.
    For k = 1 To KEPT_COUNT
        rs.Cells(1, resultCol(k - 1)).Value = names(k - 1)
    Next k
    With rs.Range(rs.Cells(1, 1), rs.Cells(1, RESULT_COLS))
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With

    If blockCount(5) > 0 Then lastBlock = 5 Else lastBlock = 4
    r = 3
    pos = 0
    For b = 1 To lastBlock
        rs.Cells(r, 1).Value = BlockName(b)
        rs.Cells(r, 1).Font.Bold = True
        rs.Cells(r, 1).Font.Size = 12
        r = r + 1
        If blockCount(b) = 0 Then
            rs.Cells(r, 1).Value = NONE_TEXT
            rs.Cells(r, 1).Font.Italic = True
            r = r + 1
        Else
            firstRow = r
            For i = 1 To blockCount(b)
                pos = pos + 1
                For k = 1 To KEPT_COUNT
                    WriteCell rs.Cells(r, resultCol(k - 1)), values(order(pos), k), formats(order(pos), k)
                Next k
                r = r + 1
            Next i
            TickBoxes rs.Range(rs.Cells(firstRow, 3), rs.Cells(r - 1, 3))
            TickBoxes rs.Range(rs.Cells(firstRow, 6), rs.Cells(r - 1, 6))
        End If
        If b < lastBlock Then r = r + 1
    Next b

    rs.Columns(1).AutoFit
    If rs.Columns(1).ColumnWidth > MAX_DESCRIPTION_WIDTH Then rs.Columns(1).ColumnWidth = MAX_DESCRIPTION_WIDTH
    rs.Columns(2).AutoFit
    rs.Columns(4).AutoFit
    rs.Columns(5).AutoFit
    rs.Columns(7).AutoFit
    rs.Columns(3).ColumnWidth = TICK_BOX_WIDTH
    rs.Columns(6).ColumnWidth = TICK_BOX_WIDTH
End Sub

' 4.4, 4.7: the value exactly as read, with the export cell's number format. Text is written
' into a text-formatted cell first so Excel cannot turn it into a number or a date, then the
' cell gets the export's format back.
Private Sub WriteCell(ByVal c As Range, ByVal v As Variant, ByVal fmt As String)
    If IsEmpty(v) Then
        c.NumberFormat = fmt
    ElseIf VarType(v) = vbString Then
        c.NumberFormat = "@"
        c.Value = v
        c.NumberFormat = fmt
    Else
        c.NumberFormat = fmt
        c.Value = v
    End If
End Sub

' A thin border on all four sides of every cell of a one-column range.
Private Sub TickBoxes(ByVal rng As Range)
    Dim edges As Variant, i As Long
    edges = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
    For i = 0 To 3
        rng.Borders(edges(i)).LineStyle = xlContinuous
        rng.Borders(edges(i)).Weight = xlThin
    Next i
    If rng.Rows.Count > 1 Then
        rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        rng.Borders(xlInsideHorizontal).Weight = xlThin
    End If
End Sub

' 4.8: A4 landscape, one page wide, heading row on every page. False when any part failed.
Private Function SetUpPrinting(ByVal rs As Worksheet) As Boolean
    Dim ok As Boolean
    ok = True
    On Error Resume Next
    With rs.PageSetup
        .PaperSize = xlPaperA4: If Err.Number <> 0 Then ok = False: Err.Clear
        .Orientation = xlLandscape: If Err.Number <> 0 Then ok = False: Err.Clear
        .Zoom = False: If Err.Number <> 0 Then ok = False: Err.Clear
        .FitToPagesWide = 1: If Err.Number <> 0 Then ok = False: Err.Clear
        .FitToPagesTall = False: If Err.Number <> 0 Then ok = False: Err.Clear
        .PrintTitleRows = "$1:$1": If Err.Number <> 0 Then ok = False: Err.Clear
        .PrintGridlines = False: If Err.Number <> 0 Then ok = False: Err.Clear
        .LeftMargin = Application.InchesToPoints(0.25): If Err.Number <> 0 Then ok = False: Err.Clear
        .RightMargin = Application.InchesToPoints(0.25): If Err.Number <> 0 Then ok = False: Err.Clear
        .TopMargin = Application.InchesToPoints(0.75): If Err.Number <> 0 Then ok = False: Err.Clear
        .BottomMargin = Application.InchesToPoints(0.75): If Err.Number <> 0 Then ok = False: Err.Clear
        .HeaderMargin = Application.InchesToPoints(0.3): If Err.Number <> 0 Then ok = False: Err.Clear
        .FooterMargin = Application.InchesToPoints(0.3): If Err.Number <> 0 Then ok = False: Err.Clear
        .CenterFooter = "Page &P of &N": If Err.Number <> 0 Then ok = False: Err.Clear
    End With
    On Error GoTo 0
    SetUpPrinting = ok
End Function

' 4.2 step 5: "Trimmed 57 vehicles into the "Trimmed" tab: Retail 23, Demo 20, Courtesy 13, Auction 0, Other 1."
Private Function FinalMessage(ByVal tabName As String, ByVal rowCount As Long, blockCount() As Long) As String
    Dim b As Long, parts As String
    For b = 1 To 4
        parts = parts & IIf(b = 1, "", ", ") & BlockName(b) & " " & blockCount(b)
    Next b
    If blockCount(5) > 0 Then parts = parts & ", " & OTHER_BLOCK & " " & blockCount(5)
    FinalMessage = "Trimmed " & rowCount & " vehicles into the """ & tabName & """ tab: " & parts & "."
End Function
