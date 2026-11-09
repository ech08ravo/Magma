Attribute VB_Name = "modViews"
' ========================================================================
'  Curriculum Map Builder - Refresh views: rebuild every sheet from Tasks
' ========================================================================
Option Explicit

Private Const BLOCK As Long = 8           ' columns per card (one per GLO)
Private Const HEAD_ROW As Long = 8        ' trimester header row on the grid sheets
Private Const AOL_TARGET As Long = 2

' Tasks helper column positions (set by BuildTasksTable)
Private tMark0 As Long, tAny0 As Long, tIp As Long, tOn As Long, tOth As Long, tWt As Long
Private tBal As Long, tAll As Long, tSort As Long, tRowId As Long, tCard As Long, tKey As Long, tVis As Long
Private tShow As Long, tLastRow As Long, nFlags As Long
' WeekRows
Private wN As Long

Public Sub RefreshViewsButton()
    Dim msg As String
    msg = RefreshViews()
    If Left$(msg, 3) = "ERR" Then MsgBox msg, vbExclamation
End Sub

Public Function RefreshViews() As String
    Dim t0 As Double, state As Collection
    t0 = Timer
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    LoadTypeCodes
    If Not ReadTasksData(Hdr, Dat, NR, NC) Then
        RefreshViews = "ERR the Tasks sheet is empty - use Import export first"
        GoTo Done
    End If
    If ColOf("Unit Code") = 0 Or ColOf("AT Number", "Task Number") = 0 Or ColOf("Study Year") = 0 Then
        RefreshViews = "ERR Tasks needs Study Year, Trimester, Unit Code and AT Number columns"
        GoTo Done
    End If
    DeriveRows
    Set state = SaveFilterState()
    DeleteViewSheets
    BuildTasksTable
    BuildWeekRows
    BuildGrid "Timeline"
    BuildGrid "Weeks"
    BuildBalance "AOL type balance", False
    BuildBalance "All tasks", True
    BuildSummary
    BuildLegend
    OrderSheets
    RestoreFilterState state
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    ThisWorkbook.Worksheets("Timeline").Activate
    ThisWorkbook.Worksheets("Timeline").Range("A1").Select
    RefreshViews = "OK " & NR & " task rows, " & Format$(Timer - t0, "0.0") & "s"
Done:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function
Fail:
    RefreshViews = "ERR " & Err.Number & " " & Err.Description & " (" & Erl & ")"
    Resume Done
End Function

' ---- sheet management ----------------------------------------------------
Private Sub DeleteViewSheets()
    Dim n As Variant, ws As Worksheet
    For Each n In Array("Timeline", "Weeks", "AOL type balance", "All tasks", "Summary", "Outline", "WeekRows", "Legend")
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(n))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
            If ThisWorkbook.Worksheets.Count > 1 Then ws.Delete
        End If
    Next n
End Sub

Private Function NewSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    ActiveWindow.DisplayGridlines = False
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = xlCenter
    Set NewSheet = ws
End Function

Private Sub OrderSheets()
    Dim order As Variant, i As Long, ws As Worksheet, n As Variant
    order = Array("Timeline", "Weeks", "AOL type balance", "All tasks", "Summary", "Tasks", "Lists", "Legend", "Outline", "WeekRows")
    For i = UBound(order) To 0 Step -1
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(order(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then ws.Move Before:=ThisWorkbook.Worksheets(1)
    Next i
    ' remove an untouched starter sheet (e.g. "Sheet1")
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Sheet#*" Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
        End If
    Next ws
    ThisWorkbook.Worksheets("Outline").Visible = xlSheetHidden
    ThisWorkbook.Worksheets("WeekRows").Visible = xlSheetHidden
End Sub

Private Function SaveFilterState() As Collection
    Dim st As New Collection, n As Variant, ws As Worksheet, a As Variant
    For Each n In Array("Timeline", "Weeks")
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(n))
        On Error GoTo 0
        If Not ws Is Nothing Then
            a = Array(ws.Range("B4:BQ5").Value)
            st.Add a, CStr(n)
        End If
    Next n
    Set SaveFilterState = st
End Function

Private Sub RestoreFilterState(ByVal st As Collection)
    Dim n As Variant, ws As Worksheet, v As Variant, g As Long, c As Long
    For Each n In Array("Timeline", "Weeks")
        If HasKey(st, CStr(n)) Then
            v = st(CStr(n))(0)
            Set ws = ThisWorkbook.Worksheets(CStr(n))
            For g = 0 To 7
                c = 1 + BLOCK * g                       ' offset within B4:BQ5 (B = 1)
                If IsNumeric(v(1, c)) Then ws.Cells(4, 2 + BLOCK * g).Value = CLng(v(1, c))
            Next g
            For g = 0 To 2
                c = 1 + BLOCK * g
                If IsNumeric(v(2, c)) Then ws.Cells(5, 2 + BLOCK * g).Value = CLng(v(2, c))
            Next g
            If IsNumeric(v(2, 1 + BLOCK * 4)) Then ws.Cells(5, 2 + BLOCK * 4).Value = v(2, 1 + BLOCK * 4)
            If IsNumeric(v(2, 1 + BLOCK * 5)) Then ws.Cells(5, 2 + BLOCK * 5).Value = v(2, 1 + BLOCK * 5)
        End If
    Next n
End Sub

' ---- small formatting helpers --------------------------------------------
Private Sub Fmt(ByVal rng As Range, Optional ByVal bold As Boolean = False, Optional ByVal size As Double = 0, _
                Optional ByVal fontHex As String = "", Optional ByVal fillHex As String = "", _
                Optional ByVal italic As Boolean = False, Optional ByVal wrap As Boolean = False, _
                Optional ByVal hAlign As Long = 0, Optional ByVal vTop As Boolean = False)
    With rng
        .Font.Bold = bold
        .Font.Italic = italic
        If size > 0 Then .Font.Size = size
        If Len(fontHex) > 0 Then .Font.Color = HexRGB(fontHex)
        If Len(fillHex) > 0 Then .Interior.Color = HexRGB(fillHex)
        .WrapText = wrap
        If hAlign <> 0 Then .HorizontalAlignment = hAlign
        If vTop Then .VerticalAlignment = xlTop
    End With
End Sub

Private Sub MergeWrite(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal value As Variant)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Merge
        .Cells(1, 1).Value = value
    End With
End Sub

Private Function Q(ByVal s As String) As String
    Q = """" & Replace(s, """", """""") & """"
End Function

' ========================================================================
'  TASKS TABLE (data + pathway columns + helpers)
' ========================================================================
Private Sub BuildTasksTable()
    Dim ws As Worksheet, paths As Collection, oneCourse As Boolean, nCols As Long, out() As Variant
    Dim r As Long, c As Long, g As Long, p As Variant, own As String, i As Long, helpers As Variant, lo As ListObject
    Set ws = EnsureSheet(TASKS_SHEET)
    ws.Cells.Clear
    Do While ws.ListObjects.Count > 0: ws.ListObjects(1).Delete: Loop
    Do While ws.Shapes.Count > 0: ws.Shapes(1).Delete: Loop
    Set paths = PathwayList()
    oneCourse = (CourseCount() <= 1)
    nFlags = paths.Count
    helpers = HelperNames()                     ' last item is "Show (Timeline)": outside the table
    nCols = NC + nFlags + UBound(helpers)       ' helpers 0..UBound-1 inside the table
    tMark0 = NC + nFlags + 1: tAny0 = tMark0 + 8
    tIp = tAny0 + 8: tOn = tIp + 1: tOth = tIp + 2: tWt = tIp + 3
    tBal = tIp + 4: tAll = tIp + 5: tSort = tIp + 6: tRowId = tIp + 7: tCard = tIp + 8: tKey = tIp + 9: tVis = tIp + 10
    tShow = nCols + 1: tLastRow = NR + 1
    ReDim out(1 To NR + 1, 1 To nCols)
    For c = 1 To NC: out(1, c) = Hdr(c): Next c
    For i = 1 To nFlags: out(1, NC + i) = UniqueHeader(PathHeader(paths(i), oneCourse), out, NC + i - 1): Next i
    For i = 0 To UBound(helpers) - 1: out(1, tMark0 + i) = helpers(i): Next i
    Dim cPath As Long: cPath = ColOf("Pathways")
    For r = 1 To NR
        For c = 1 To NC: out(r + 1, c) = Dat(r, c): Next c
        If cPath > 0 Then own = ";" & Replace(Dat(r, cPath), "; ", ";") & ";" Else own = ""
        For i = 1 To nFlags
            If InStr(own, ";" & paths(i) & ";") > 0 Then out(r + 1, NC + i) = GDot() Else out(r + 1, NC + i) = ""
        Next i
        For g = 1 To 8
            out(r + 1, tMark0 + g - 1) = RMarks(r, g)
            out(r + 1, tAny0 + g - 1) = RAny(r, g)
        Next g
        out(r + 1, tIp) = RIp(r): out(r + 1, tOn) = ROn(r): out(r + 1, tOth) = ROth(r): out(r + 1, tWt) = RWt(r)
        out(r + 1, tBal) = RBal(r): out(r + 1, tAll) = RAll(r): out(r + 1, tSort) = RSort(r)
        out(r + 1, tRowId) = r: out(r + 1, tCard) = RCard(r): out(r + 1, tKey) = RKey(r)
    Next r
    ws.Range(ws.Cells(1, 1), ws.Cells(NR + 1, NC + nFlags)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(NR + 1, nCols)).Value = out
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(NR + 1, nCols)), , xlYes)
    lo.Name = "Tasks"
    lo.TableStyle = "TableStyleLight1"
    ' Visible: 1 while the row is shown, 0 once a filter hides it
    ws.Range(ws.Cells(2, tVis), ws.Cells(NR + 1, tVis)).Formula = "=SUBTOTAL(103," & ColLetter(ColOf("Unit Code")) & "2)"
    ws.Cells(1, tShow).Value = "Show (Timeline)"
    ' look
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))
        .Font.Bold = True: .WrapText = True: .VerticalAlignment = xlBottom
        .Interior.Color = HexRGB("#2B2F3A"): .Font.Color = vbWhite
    End With
    If nFlags > 0 Then
        With ws.Range(ws.Cells(1, NC + 1), ws.Cells(1, NC + nFlags))
            .Orientation = 90: .HorizontalAlignment = xlCenter: .Interior.Color = HexRGB("#3B4252")
        End With
        With ws.Range(ws.Cells(2, NC + 1), ws.Cells(NR + 1, NC + nFlags))
            .HorizontalAlignment = xlCenter: .Font.Bold = True: .Font.Color = HexRGB("#1F6F66")
        End With
        ws.Range(ws.Cells(1, NC + 1), ws.Cells(1, NC + nFlags)).EntireColumn.ColumnWidth = 3.2
        ws.Rows(1).RowHeight = 150
    Else
        ws.Rows(1).RowHeight = 30
    End If
    ws.Range(ws.Cells(1, 1), ws.Cells(NR + 1, NC)).ColumnWidth = 12
    SetWidth ws, "Study Year", 6: SetWidth ws, "Trimester", 11: SetWidth ws, "AT Group", 7
    SetWidth ws, "Unit Code", 9: SetWidth ws, "AT Number", 10: SetWidth ws, "Assessment Title", 34
    SetWidth ws, "Course", 22: SetWidth ws, "Pathways", 30: SetWidth ws, "Student Output", 30
    SetWidth ws, "AoL Task Type", 22: SetWidth ws, "Weighting", 18: SetWidth ws, "Guide Evidence", 30
    SetWidth ws, "Tools", 16: SetWidth ws, "Assessment type", 18: SetWidth ws, "Planned Weeks", 10
    ws.Range(ws.Cells(1, tMark0), ws.Cells(1, tShow)).EntireColumn.Hidden = True
    ws.Range(ws.Cells(2, 1), ws.Cells(NR + 1, nCols)).RowHeight = 15
    ws.Range(ws.Cells(2, 1), ws.Cells(NR + 1, NC)).WrapText = False
    ' GLO state colouring on the export's GLO columns
    Dim gc As Variant: gc = GloCols()
    For g = 1 To 8
        If gc(g) > 0 Then
            With ws.Range(ws.Cells(2, gc(g)), ws.Cells(NR + 1, gc(g))).FormatConditions
                .Add(xlCellValue, xlEqual, "=""Yes""").Interior.Color = HexRGB("#1F6F66")
                .Item(.Count).Font.Color = vbWhite
                .Add(xlCellValue, xlEqual, "=""Non-hurdle Yes""").Interior.Color = HexRGB("#BFE3DE")
                .Add(xlCellValue, xlEqual, "=""Data Error""").Interior.Color = HexRGB("#E05252")
                .Add(xlCellValue, xlEqual, "=""Not Assessed""").Font.Color = HexRGB("#9AA0A6")
                .Add(xlCellValue, xlEqual, "=""No""").Font.Color = HexRGB("#9AA0A6")
            End With
        End If
    Next g
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(2, ColOf("AT Number", "Task Number") + 1).Select
    ActiveWindow.FreezePanes = True
    AddButtons ws, ws.Cells(1, 1).Left + 4, 4, False
End Sub

Private Function UniqueHeader(ByVal h As String, out() As Variant, ByVal upto As Long) As String
    Dim c As Long, again As Boolean
    Do
        again = False
        For c = 1 To upto
            If CStr(out(1, c)) = h Then h = h & " ": again = True: Exit For
        Next c
    Loop While again
    UniqueHeader = h
End Function

Private Sub SetWidth(ByVal ws As Worksheet, ByVal colName As String, ByVal w As Double)
    Dim c As Long: c = ColOf(colName)
    If c > 0 Then ws.Columns(c).ColumnWidth = w
End Sub

Private Function TRange(ByVal c As Long) As String
    TRange = ColRange(TASKS_SHEET, c, tLastRow)
End Function

' ========================================================================
'  WEEKROWS (hidden): one row per task per week
' ========================================================================
Private Sub BuildWeekRows()
    Dim ws As Worksheet, rows As New Collection, r As Long, wk As Variant, wkLine As String, detail As String
    Dim planned As Collection, guide As Collection, n As Long, out() As Variant, i As Long, g As Long, item As Variant
    Dim multi As Boolean, wtxt As String, w As Double, seen As Collection, atn As Long
    multi = (CourseCount() > 1)
    For r = 1 To NR
        w = ShortWeight(V(r, "Weighting"))
        If w >= 0 Then wtxt = Format$(w * 100, "0.##") & "%" Else wtxt = GMdash()
        wtxt = Replace(wtxt, ".%", "%")
        Set planned = Numbers(V(r, "Planned Weeks"))
        Set guide = Numbers(V(r, "Guide Weeks"))
        Set seen = New Collection
        If planned.Count > 0 Then
            For Each wk In planned
                If Not HasKey(seen, "k" & wk) Then
                    seen.Add wk, "k" & wk
                    rows.Add Array(r, CStr(wk), CardText(r, "Planned: Wk " & wk, multi, RHasCalmGlo(r), wtxt))
                End If
            Next wk
        ElseIf guide.Count > 0 Then
            detail = V(r, "Guide Week Detail")
            For Each wk In guide
                If Not HasKey(seen, "k" & wk) Then
                    seen.Add wk, "k" & wk
                    wkLine = "Guide: Wk " & wk & GuidePart(detail, CLng(wk))
                    rows.Add Array(r, CStr(wk), CardText(r, wkLine, multi, RHasCalmGlo(r), wtxt))
                End If
            Next wk
        Else
            atn = FirstNumber(V(r, "AT Number"))
            If atn >= 0 Then
                rows.Add Array(r, CStr(atn), CardText(r, FallbackLine(), multi, RHasCalmGlo(r), wtxt))
            Else
                rows.Add Array(r, "none", CardText(r, "No week", multi, RHasCalmGlo(r), wtxt))
            End If
        End If
    Next r
    wN = rows.Count
    Set ws = NewSheet("WeekRows")
    ReDim out(1 To wN + 1, 1 To 18)
    Dim hdrs As Variant: hdrs = Array("Row ID", "Study Year", "Trimester", "Week", "Card", "Visible", _
        "GLO1 Mark", "GLO2 Mark", "GLO3 Mark", "GLO4 Mark", "GLO5 Mark", "GLO6 Mark", "GLO7 Mark", "GLO8 Mark", _
        "In person", "Online", "Other mode", "Weight %")
    For i = 0 To 17: out(1, i + 1) = hdrs(i): Next i
    For i = 1 To wN
        item = rows(i): r = item(0)
        out(i + 1, 1) = r: out(i + 1, 2) = V(r, "Study Year"): out(i + 1, 3) = V(r, "Trimester")
        out(i + 1, 4) = item(1): out(i + 1, 5) = item(2)
        For g = 1 To 8: out(i + 1, 6 + g) = RMarks(r, g): Next g
        out(i + 1, 15) = RIp(r): out(i + 1, 16) = ROn(r): out(i + 1, 17) = ROth(r): out(i + 1, 18) = RWt(r)
    Next i
    ws.Range("D:D").NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(wN + 1, 18)).Value = out
    ws.Range(ws.Cells(2, 6), ws.Cells(wN + 1, 6)).Formula = _
        "=IFERROR(INDEX(" & TRange(tVis) & ",MATCH(A2," & TRange(tRowId) & ",0)),0)"
    ws.Cells(1, 19).Value = "Show (Weeks)"
End Sub

Private Function GuidePart(ByVal detail As String, ByVal wk As Long) As String
    Dim key As String, p As Long, e As Long
    key = "Wk " & wk & " ("
    p = InStr(detail, key)
    If p > 0 Then
        e = InStr(p + Len(key), detail, ")")
        If e > 0 Then GuidePart = " (" & Mid$(detail, p + Len(key), e - p - Len(key)) & ")"
    End If
End Function

' ========================================================================
'  TIMELINE / WEEKS grids
' ========================================================================
Private Sub BuildGrid(ByVal sheetName As String)
    Dim ws As Worksheet, isWeeks As Boolean, srcSheet As String, srcLast As Long
    Dim yearR As String, triR As String, keyR As String, cardR As String, marksR As String, showR As String
    Dim years As Variant, tris As Variant, keys() As Variant, nKeys() As Long, t As Long, k As Long
    Dim slotsTri() As String, slotsKey() As String, nSlots As Long, j As Long, y As Long, d As Long, depth As Long
    Dim row As Long, c As Long, crit As String, g As Long, r As Long, title As String
    isWeeks = (sheetName = "Weeks")
    Set ws = NewSheet(sheetName)
    years = YearList(): tris = TrimesterList()
    ' source ranges
    If isWeeks Then
        srcSheet = "WeekRows": srcLast = wN + 1
        yearR = ColRange(srcSheet, 2, srcLast): triR = ColRange(srcSheet, 3, srcLast)
        keyR = ColRange(srcSheet, 4, srcLast): cardR = ColRange(srcSheet, 5, srcLast)
        marksR = "WeekRows!$G$2:$N$" & srcLast
        showR = ColRange(srcSheet, 19, srcLast)
    Else
        srcSheet = TASKS_SHEET: srcLast = tLastRow
        yearR = TRange(ColOf("Study Year")): triR = TRange(ColOf("Trimester"))
        keyR = TRange(ColOf("AT Group")): cardR = TRange(tCard)
        marksR = TASKS_SHEET & "!$" & ColLetter(tMark0) & "$2:$" & ColLetter(tMark0 + 7) & "$" & tLastRow
        showR = TRange(tShow)
    End If
    ' columns: keys per trimester (AT groups, or weeks found in that trimester)
    ReDim keys(LBound(tris) To UBound(tris)): ReDim nKeys(LBound(tris) To UBound(tris))
    For t = LBound(tris) To UBound(tris)
        If isWeeks Then keys(t) = WeekKeysFor(CStr(tris(t))) Else keys(t) = Array("AT1", "AT2", "AT3", "Other")
        nKeys(t) = UBound(keys(t)) - LBound(keys(t)) + 1
        If nKeys(t) = 0 Then nKeys(t) = 0
    Next t
    nSlots = 0
    For t = LBound(tris) To UBound(tris): nSlots = nSlots + nKeys(t): Next t
    If nSlots = 0 Then nSlots = 1
    ReDim slotsTri(1 To nSlots): ReDim slotsKey(1 To nSlots)
    j = 0
    For t = LBound(tris) To UBound(tris)
        For k = LBound(keys(t)) To UBound(keys(t))
            j = j + 1: slotsTri(j) = tris(t): slotsKey(j) = keys(t)(k)
        Next k
    Next t
    ' --- per-row Show helper: every filter worked out once per source row ---
    WriteShowFormulas ws, isWeeks
    Dim boxes(1 To 8) As String, ipB As String, onB As String, othB As String, wMin As String, wMax As String
    For g = 1 To 8: boxes(g) = "$" & ColLetter(2 + BLOCK * (g - 1)) & "$4": Next g
    ipB = "$B$5": onB = "$" & ColLetter(2 + BLOCK) & "$5": othB = "$" & ColLetter(2 + 2 * BLOCK) & "$5"
    wMin = "$" & ColLetter(2 + 4 * BLOCK) & "$5": wMax = "$" & ColLetter(2 + 5 * BLOCK) & "$5"
    ' --- header area ---
    title = CourseTitle()
    ws.Range("A1").Value = title: Fmt ws.Range("A1"), True, 14
    Dim ticked As String, which As String
    ticked = "(" & Join(boxes, "+") & ")"
    which = "TEXTJOIN("", "",TRUE"
    For g = 1 To 8: which = which & ",IF(" & boxes(g) & "=1,""GLO" & g & ""","""")": Next g
    which = which & ")"
    ws.Range("A2").Formula2 = "=" & Q(IIf(isWeeks, "Assessment weeks", "Assessment timeline") & SepDot()) & _
        "&SUMPRODUCT(--(" & showR & "=1))&" & Q(" cards shown" & SepDot()) & _
        "&IF(" & ticked & "=0," & Q("click a GLO below to show only its tasks") & "," & Q("GLO filter: ") & "&" & which & "&" & Q(" (click again to clear)") & ")" & _
        "&IF(" & ipB & "+" & onB & "+" & othB & "=0,""""," & Q(SepDot() & "Mode: ") & "&TEXTJOIN("" or "",TRUE,IF(" & ipB & "=1,""in person"",""""),IF(" & onB & "=1,""online"",""""),IF(" & othB & "=1,""other"","""")))" & _
        "&IF(AND(" & wMin & "<=0," & wMax & ">=100),""""," & Q(SepDot() & "Weight ") & "&" & wMin & "&" & Q(GNdash()) & "&" & wMax & "&""%"")" & _
        "&" & Q(SepDot() & "Tasks-sheet filters also apply")
    Fmt ws.Range("A2"), False, 0, "#6B7280", "", True
    ' row 3: alignment legend; row 4: GLO toggles; row 5: mode toggles + weighting; row 6: note
    Dim mk As Variant
    c = 2
    For Each mk In AllMarks()
        MergeWrite ws, 3, c, c + 7, mk & " " & MarkLabel(CStr(mk))
        Fmt ws.Range(ws.Cells(3, c), ws.Cells(3, c + 7)), False, 9, "", MarkFill(CStr(mk))
        c = c + 8
    Next mk
    If isWeeks Then
        MergeWrite ws, 3, c, c + 15, FallbackLine() & " (no guide or planned week): dashed border"
        Fmt ws.Range(ws.Cells(3, c), ws.Cells(3, c + 15)), False, 9
        ws.Range(ws.Cells(3, c), ws.Cells(3, c + 15)).Borders.LineStyle = xlDash
    End If
    For g = 1 To 8
        c = 2 + BLOCK * (g - 1)
        ToggleCell ws.Cells(4, c)
        MergeWrite ws, 4, c + 1, c + 7, GDot() & " GLO" & g & " " & GloName(g)
        Fmt ws.Range(ws.Cells(4, c + 1), ws.Cells(4, c + 7)), True, 9, GloColour(g)
        With ws.Range(ws.Cells(4, c + 1), ws.Cells(4, c + 7)).FormatConditions.Add(xlExpression, , "=" & boxes(g) & "=1")
            .Interior.Color = HexRGB("#EEF0F3"): .Font.Underline = xlUnderlineStyleSingle
        End With
    Next g
    ws.Rows(4).RowHeight = 18
    ws.Range("A5").Value = "Filter": Fmt ws.Range("A5"), True, 9, "#6B7280"
    Dim labels As Variant, modeBoxes As Variant
    labels = Array("In person (campus, site, field, placement)", "Online", "Other / unspecified mode")
    modeBoxes = Array(ipB, onB, othB)
    For g = 0 To 2
        c = 2 + BLOCK * g
        ToggleCell ws.Cells(5, c)
        MergeWrite ws, 5, c + 1, c + 7, labels(g)
        Fmt ws.Range(ws.Cells(5, c + 1), ws.Cells(5, c + 7)), True, 9, "#303842"
        With ws.Range(ws.Cells(5, c + 1), ws.Cells(5, c + 7)).FormatConditions.Add(xlExpression, , "=" & modeBoxes(g) & "=1")
            .Interior.Color = HexRGB("#EEF0F3"): .Font.Underline = xlUnderlineStyleSingle
        End With
    Next g
    MergeWrite ws, 5, 2 + 3 * BLOCK, 1 + 4 * BLOCK, "Weighting from (%)"
    Fmt ws.Range(ws.Cells(5, 2 + 3 * BLOCK), ws.Cells(5, 1 + 4 * BLOCK)), True, 9, "#303842", "", False, False, xlRight
    MergeWrite ws, 5, 2 + 4 * BLOCK, 4 + 4 * BLOCK, 0
    MergeWrite ws, 5, 5 + 4 * BLOCK, 1 + 5 * BLOCK, "to"
    MergeWrite ws, 5, 2 + 5 * BLOCK, 4 + 5 * BLOCK, 100
    For Each mk In Array(2 + 4 * BLOCK, 2 + 5 * BLOCK)
        Fmt ws.Range(ws.Cells(5, mk), ws.Cells(5, mk + 2)), True, 10, "", "#FFF8DC", False, False, xlCenter
        ws.Range(ws.Cells(5, mk), ws.Cells(5, mk + 2)).Borders.Color = HexRGB("#C8CDD3")
        With ws.Cells(5, mk).Validation
            .Delete
            .Add xlValidateDecimal, xlValidAlertStop, xlBetween, "0", "100"
            .ErrorMessage = "Enter a percentage from 0 to 100."
        End With
    Next mk
    MergeWrite ws, 5, 5 + 5 * BLOCK, 1 + 6 * BLOCK, "%  (unreadable weights show only at 0" & GNdash() & "100)"
    Fmt ws.Range(ws.Cells(5, 5 + 5 * BLOCK), ws.Cells(5, 1 + 6 * BLOCK)), False, 9, "#6B7280"
    ws.Rows(5).RowHeight = 18
    ws.Range("B6").Value = IIf(isWeeks, _
        "Planned Weeks first, then unit-guide weeks, then the AT number (AT1 " & GArrow() & " Week 1 " & GEll() & ", dashed, not a due date). A task with several weeks appears in each.", _
        "Tasks are grouped by AT number within each unit teaching trimester (not due dates). Units taught in several trimesters appear in each.") & _
        " Click a GLO or mode label to filter. Under each card, one cell per GLO (GLO1" & GArrow() & "GLO8) shows the task's mark in that GLO's colour."
    Fmt ws.Range("B6"), False, 9, "#6B7280"
    ' column widths
    ws.Columns(1).ColumnWidth = 7
    ws.Range(ws.Cells(1, 2), ws.Cells(1, 1 + nSlots * BLOCK)).EntireColumn.ColumnWidth = 2.9
    ' trimester header row with totals
    c = 2
    For t = LBound(tris) To UBound(tris)
        If nKeys(t) > 0 Then
            ws.Range(ws.Cells(HEAD_ROW, c), ws.Cells(HEAD_ROW, c + nKeys(t) * BLOCK - 1)).Merge
            ws.Cells(HEAD_ROW, c).Formula = HeaderFormula(CStr(tris(t)), "(" & triR & "=" & Q(CStr(tris(t))) & ")*(" & showR & "=1)", cardR, " " & GMid() & " ", " cards")
            Fmt ws.Range(ws.Cells(HEAD_ROW, c), ws.Cells(HEAD_ROW, c + nKeys(t) * BLOCK - 1)), True, 10, "#303842", "#EDF0F4"
            c = c + nKeys(t) * BLOCK
        End If
    Next t
    ws.Cells(HEAD_ROW, 1).Value = "Year": Fmt ws.Cells(HEAD_ROW, 1), True, 10, "#303842", "#EDF0F4"
    ' bands
    row = HEAD_ROW + 1
    Dim bandStart As Long: bandStart = row
    For y = LBound(years) To UBound(years)
        depth = 1
        For j = 1 To nSlots
            d = SlotCount(isWeeks, CStr(years(y)), slotsTri(j), slotsKey(j))
            If d > depth Then depth = d
        Next j
        ws.Cells(row, 1).Value = years(y)
        Fmt ws.Cells(row, 1), True, 12, "#303842", "", False, False, 0, True
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 1 + nSlots * BLOCK)).Borders(xlEdgeTop).Color = HexRGB("#6B7280")
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 1 + nSlots * BLOCK)).Borders(xlEdgeTop).Weight = xlMedium
        For j = 1 To nSlots
            c = 2 + (j - 1) * BLOCK
            ws.Range(ws.Cells(row, c), ws.Cells(row, c + BLOCK - 1)).Merge
            crit = "(" & yearR & "=" & Q(CStr(years(y))) & ")*(" & triR & "=" & Q(slotsTri(j)) & ")*(" & keyR & "=" & Q(slotsKey(j)) & ")*(" & showR & "=1)"
            ws.Cells(row, c).Formula = HeaderFormula(KeyLabel(isWeeks, slotsKey(j)), crit, cardR, " (", ")")
            Fmt ws.Range(ws.Cells(row, c), ws.Cells(row, c + BLOCK - 1)), True, 9, "#59636F"
            ws.Range(ws.Cells(row, c), ws.Cells(row, c + BLOCK - 1)).Borders(xlEdgeBottom).Color = HexRGB(KeyRule(isWeeks, slotsKey(j)))
            ws.Range(ws.Cells(row, c), ws.Cells(row, c + BLOCK - 1)).Borders(xlEdgeBottom).Weight = xlMedium
        Next j
        ws.Rows(row).RowHeight = 16
        r = row + 1
        ' strip colours per GLO column, then cards on top
        For j = 1 To nSlots
            c = 2 + (j - 1) * BLOCK
            For g = 1 To 8
                With ws.Range(ws.Cells(r, c + g - 1), ws.Cells(r + 2 * depth - 1, c + g - 1))
                    .Font.Color = HexRGB(GloColour(g)): .Font.Bold = True: .HorizontalAlignment = xlCenter
                End With
            Next g
        Next j
        For d = 1 To depth
            ws.Rows(r).RowHeight = 54: ws.Rows(r + 1).RowHeight = 14
            ws.Cells(r, 1).Value = "c": ws.Cells(r + 1, 1).Value = "s"
            ws.Range(ws.Cells(r, 1), ws.Cells(r + 1, 1)).Font.Color = vbWhite
            For j = 1 To nSlots
                c = 2 + (j - 1) * BLOCK
                crit = "(" & yearR & "=" & Q(CStr(years(y))) & ")*(" & triR & "=" & Q(slotsTri(j)) & ")*(" & keyR & "=" & Q(slotsKey(j)) & ")*(" & showR & "=1)"
                With ws.Range(ws.Cells(r, c), ws.Cells(r, c + BLOCK - 1))
                    .Merge
                    .WrapText = True: .VerticalAlignment = xlTop: .HorizontalAlignment = xlLeft
                    .Font.Color = HexRGB("#20242C"): .Font.Bold = False: .Font.Size = 9
                End With
                ws.Cells(r, c).Formula2 = "=IFERROR(INDEX(FILTER(" & cardR & "," & crit & ")," & d & "),"""")"
                ' GLO strip: one formula spills the 8 marks of the d-th card
                ws.Cells(r + 1, c).Formula2 = "=IFERROR(INDEX(FILTER(" & marksR & "," & crit & ")," & d & ",0)&"""","""")"
            Next j
            r = r + 2
        Next d
        row = r
    Next y
    ' fills keyed off the card's first character (the mark); strip rows look one row up
    Dim area As Range, fc As FormatCondition
    Set area = ws.Range(ws.Cells(bandStart, 2), ws.Cells(row, 1 + nSlots * BLOCK))
    For Each mk In AllMarks()
        Set fc = area.FormatConditions.Add(xlExpression, , _
            "=LEFT(INDEX($1:$1048576,ROW()-($A" & bandStart & "=""s""),2+8*INT((COLUMN()-2)/8)),1)=" & Q(CStr(mk)))
        fc.Interior.Color = HexRGB(MarkFill(CStr(mk)))
    Next mk
    If isWeeks Then
        Set fc = area.FormatConditions.Add(xlTextString, String:=FallbackLine(), TextOperator:=xlContains)
        fc.Borders.LineStyle = xlDash
        fc.Font.Color = HexRGB("#4B5563")
    End If
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(HEAD_ROW + 1, 2).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    If Not isWeeks Then AddButtons ws, ws.Cells(1, 2 + 6 * BLOCK).Left, 2, True
    With ws.PageSetup
        .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintTitleRows = "$" & HEAD_ROW & ":$" & HEAD_ROW
    End With
End Sub

Private Sub ToggleCell(ByVal cell As Range)
    cell.Value = 0
    cell.NumberFormat = """" & GBoxOn() & """;;""" & GBoxOff() & """"
    cell.HorizontalAlignment = xlCenter
    cell.Font.Size = 12
End Sub

Private Function HeaderFormula(ByVal label As String, ByVal base As String, ByVal cardR As String, _
                               ByVal sep As String, ByVal unitText As String) As String
    Dim total As String, marks As String, mk As Variant
    total = "SUMPRODUCT(" & base & ")"
    For Each mk In Array(GFlag(), GDot(), GDiamond())
        marks = marks & "&" & Q("  " & mk) & "&SUMPRODUCT(" & base & "*(LEFT(" & cardR & ",1)=" & Q(CStr(mk)) & "))"
    Next mk
    HeaderFormula = "=" & Q(label & sep) & "&" & total & "&" & Q(unitText) & "&IF(" & total & "=0,""""," & Mid$(marks, 2) & ")"
End Function

Private Sub WriteShowFormulas(ByVal grid As Worksheet, ByVal isWeeks As Boolean)
    Dim src As Worksheet, here As String, f As String, g As Long, n As Long, colShow As Long
    Dim cVis As String, cM(1 To 8) As String, cIp As String, cOn As String, cOth As String, cWt As String
    Dim boxes(1 To 8) As String, glo As String, glo2 As String
    here = "'" & grid.Name & "'!"
    If isWeeks Then
        Set src = ThisWorkbook.Worksheets("WeekRows"): n = wN: colShow = 19
        cVis = "F2": For g = 1 To 8: cM(g) = ColLetter(6 + g) & "2": Next g
        cIp = "O2": cOn = "P2": cOth = "Q2": cWt = "R2"
    Else
        Set src = ThisWorkbook.Worksheets(TASKS_SHEET): n = NR: colShow = tShow
        cVis = ColLetter(tVis) & "2": For g = 1 To 8: cM(g) = ColLetter(tMark0 + g - 1) & "2": Next g
        cIp = ColLetter(tIp) & "2": cOn = ColLetter(tOn) & "2": cOth = ColLetter(tOth) & "2": cWt = ColLetter(tWt) & "2"
    End If
    For g = 1 To 8: boxes(g) = here & "$" & ColLetter(2 + BLOCK * (g - 1)) & "$4": Next g
    glo = "OR((" & Join(boxes, "+") & ")=0,("
    For g = 1 To 8: glo = glo & IIf(g > 1, "+", "") & "(" & cM(g) & "<>"""")*" & boxes(g): Next g
    glo = glo & ")>0)"
    Dim ipB As String, onB As String, othB As String, wMin As String, wMax As String
    ipB = here & "$B$5": onB = here & "$" & ColLetter(2 + BLOCK) & "$5": othB = here & "$" & ColLetter(2 + 2 * BLOCK) & "$5"
    wMin = here & "$" & ColLetter(2 + 4 * BLOCK) & "$5": wMax = here & "$" & ColLetter(2 + 5 * BLOCK) & "$5"
    f = "=IF(AND(" & cVis & "=1," & glo & _
        ",OR((" & ipB & "+" & onB & "+" & othB & ")=0,(" & cIp & "*" & ipB & "+" & cOn & "*" & onB & "+" & cOth & "*" & othB & ")>0)" & _
        ",OR(AND(" & wMin & "<=0," & wMax & ">=100),AND(" & cWt & "<>""""," & cWt & ">=" & wMin & "," & cWt & "<=" & wMax & "))),1,0)"
    src.Range(src.Cells(2, colShow), src.Cells(n + 1, colShow)).Formula = f
End Sub

Private Function KeyLabel(ByVal isWeeks As Boolean, ByVal k As String) As String
    If Not isWeeks Then KeyLabel = k: Exit Function
    If k = "none" Then KeyLabel = "No week" Else KeyLabel = "Week " & k
End Function

Private Function KeyRule(ByVal isWeeks As Boolean, ByVal k As String) As String
    If isWeeks Then KeyRule = "#8A929C": Exit Function
    Select Case k
        Case "AT1": KeyRule = "#5277AD"
        Case "AT2": KeyRule = "#A56D2B"
        Case "AT3": KeyRule = "#56816D"
        Case Else: KeyRule = "#8A929C"
    End Select
End Function

' rows in one slot (sets band depth); unfiltered
Private Function SlotCount(ByVal isWeeks As Boolean, ByVal y As String, ByVal t As String, ByVal k As String) As Long
    Dim r As Long, n As Long, ws As Worksheet, vv As Variant
    If isWeeks Then
        Set ws = ThisWorkbook.Worksheets("WeekRows")
        vv = ws.Range(ws.Cells(2, 2), ws.Cells(wN + 1, 4)).Value
        For r = 1 To wN
            If CStr(vv(r, 1)) = y And CStr(vv(r, 2)) = t And CStr(vv(r, 3)) = k Then n = n + 1
        Next r
    Else
        Dim grp As String
        For r = 1 To NR
            grp = V(r, "AT Group")
            If grp <> "AT1" And grp <> "AT2" And grp <> "AT3" Then grp = "Other"
            If V(r, "Study Year") = y And V(r, "Trimester") = t And grp = k Then n = n + 1
        Next r
    End If
    SlotCount = n
End Function

Private Function WeekKeysFor(ByVal tri As String) As Variant
    Dim ws As Worksheet, v As Variant, r As Long, seen As New Collection, ks() As Variant, sk() As Variant, i As Long
    Set ws = ThisWorkbook.Worksheets("WeekRows")
    v = ws.Range(ws.Cells(2, 3), ws.Cells(wN + 1, 4)).Value
    For r = 1 To wN
        If CStr(v(r, 1)) = tri Then AddUnique seen, CStr(v(r, 2))
    Next r
    If seen.Count = 0 Then WeekKeysFor = Array(): Exit Function
    ReDim ks(1 To seen.Count): ReDim sk(1 To seen.Count)
    For i = 1 To seen.Count
        ks(i) = seen(i)
        If ks(i) = "none" Then sk(i) = "9999" Else sk(i) = Format$(CLng(ks(i)), "0000")
    Next i
    SortByKeys ks, sk
    WeekKeysFor = ks
End Function

Public Function YearList() As Variant
    Dim seen As New Collection, r As Long, ys() As Variant, ks() As Variant, i As Long
    For r = 1 To NR: AddUnique seen, V(r, "Study Year"): Next r
    ReDim ys(1 To seen.Count): ReDim ks(1 To seen.Count)
    For i = 1 To seen.Count: ys(i) = seen(i): ks(i) = YearSortKey(CStr(ys(i))): Next i
    SortByKeys ys, ks
    YearList = ys
End Function

Public Function TrimesterList() As Variant
    Dim seen As New Collection, r As Long, ts() As Variant, ks() As Variant, i As Long
    For r = 1 To NR: AddUnique seen, V(r, "Trimester"): Next r
    ReDim ts(1 To seen.Count): ReDim ks(1 To seen.Count)
    For i = 1 To seen.Count: ts(i) = seen(i): ks(i) = TriSortKey(CStr(ts(i))): Next i
    SortByKeys ts, ks
    TrimesterList = ts
End Function

Public Function CourseTitle() As String
    Dim seen As New Collection, r As Long, i As Long, s As String
    For r = 1 To NR
        If Len(V(r, "Course")) > 0 Then AddUnique seen, V(r, "Course")
    Next r
    For i = 1 To seen.Count: s = s & IIf(i > 1, SepDot(), "") & seen(i): Next i
    If Len(s) = 0 Then s = "Curriculum map"
    CourseTitle = s
End Function

' ========================================================================
'  AOL TYPE BALANCE / ALL TASKS
' ========================================================================
Private Sub BuildBalance(ByVal sheetName As String, ByVal allMode As Boolean)
    Dim ws As Worksheet, years As Variant, y As Long, g As Long, sc As Long, dc As Long, row As Long, head As Long
    Dim yearR As String, visR As String, lineR As String, sortR As String, mR As String, base As String, keep As String
    Dim lst As String, statusCells As String, f As String, maxLines As Long, k As Long
    Set ws = NewSheet(sheetName)
    years = YearList()
    yearR = TRange(ColOf("Study Year")): visR = TRange(tVis): sortR = TRange(tSort)
    lineR = TRange(IIf(allMode, tAll, tBal))
    If allMode Then keep = "" Else keep = "*((LEFT(" & lineR & ",1)=" & Q(GFlag()) & ")+(LEFT(" & lineR & ",1)=" & Q(GDot()) & "))"
    ws.Range("A1").Value = CourseTitle(): Fmt ws.Range("A1"), True, 14
    If allMode Then
        ws.Range("A3").Value = "Every task under each GLO it addresses, by year level: " & GFlag() & " AoL + hurdle, " & GDot() & " AoL, " & _
            GDiamond() & " hurdle only, " & GRing() & " neither, in the GLO's colour. GLO alignment from unit guides where the export has it; " & _
            "otherwise CALM, which only aligns AoL tasks (lines marked " & GMid() & " CALM). Type* = CALM Student Output until Assessment type is filled in."
    Else
        ws.Range("A3").Value = "Target: every GLO assessed in at least " & AOL_TARGET & " AoL tasks (" & GFlag() & " AoL + hurdle, " & GDot() & _
            " AoL) at every year level (CALM records GLO alignments on AoL tasks only). Type* = CALM Student Output until Assessment type is filled in. " & _
            "Follows filters on the Tasks sheet (e.g. one pathway's " & GDot() & " column = one major)."
    End If
    Fmt ws.Range("A3"), False, 9, "#4B5563"
    Dim legend As String
    For k = 1 To TCount: legend = legend & IIf(k > 1, SepDot(), "") & TCode(k) & " " & TLabel(k): Next k
    ws.Range("A4").Value = "Task types: " & legend & "  " & GMdash() & " codes under the marks are in the same order as the task list (edit the list on the Lists sheet)."
    Fmt ws.Range("A4"), False, 9, "#4B5563"
    head = 6
    ws.Columns(1).ColumnWidth = 24
    ws.Cells(head, 1).Value = "GLO"
    Fmt ws.Cells(head, 1), True, 10, "#303842", "#EDF0F4", False, False, xlCenter
    For y = LBound(years) To UBound(years)
        sc = 2 + 2 * (y - LBound(years)): dc = sc + 1
        ws.Columns(sc).ColumnWidth = 10: ws.Columns(dc).ColumnWidth = 44
        MergeWrite ws, head, sc, dc, YearLevelLabel(CStr(years(y)))
        Fmt ws.Range(ws.Cells(head, sc), ws.Cells(head, dc)), True, 10, "#303842", "#EDF0F4", False, False, xlCenter
    Next y
    ' type mix row (values filled from hidden helper rows below)
    ws.Cells(head + 1, 1).Value = "Type mix (year)"
    Fmt ws.Range(ws.Cells(head + 1, 1), ws.Cells(head + 1, 1 + 2 * (UBound(years) - LBound(years) + 1))), True, 10, "#303842", "#F7F8FA", False, False, xlCenter
    ws.Rows(head + 1).RowHeight = 22
    row = head + 2
    For g = 1 To 8
        mR = TRange(tAny0 + g - 1)
        If Not allMode Then mR = TRange(tMark0 + g - 1)
        ws.Range(ws.Cells(row, 1), ws.Cells(row + 1, 1)).Merge
        ws.Cells(row, 1).Value = GDot() & " GLO" & g & vbLf & GloName(g)
        Fmt ws.Range(ws.Cells(row, 1), ws.Cells(row + 1, 1)), True, 10, GloColour(g), "", False, True, 0, True
        maxLines = 1
        For y = LBound(years) To UBound(years)
            sc = 2 + 2 * (y - LBound(years)): dc = sc + 1
            base = "(" & yearR & "=" & Q(CStr(years(y))) & ")*(" & visR & "=1)*(" & mR & "<>"""")"
            lst = "SORTBY(FILTER(" & lineR & "," & base & keep & "),FILTER(" & sortR & "," & base & keep & "))"
            ws.Cells(row, sc).Formula2 = "=IFERROR(ROWS(UNIQUE(" & lst & ")),0)"
            ws.Cells(row, sc).NumberFormat = IIf(allMode, "0"" tasks""", """AoL ""0"" of " & AOL_TARGET & """")
            Fmt ws.Cells(row, sc), True, 10, "", "", False, False, xlCenter, True
            statusCells = statusCells & "," & ws.Cells(row, sc).Address(False, False)
            ws.Cells(row, dc).Formula2 = "=IFERROR(CONCAT(LEFT(UNIQUE(" & lst & "),1))&CHAR(10)&TEXTJOIN("" "",TRUE,TEXTAFTER(UNIQUE(" & lst & ")," & Q(GSect()) & ",-1)),"""")"
            Fmt ws.Cells(row, dc), True, 13, GloColour(g), "", False, True, 0, True
            ws.Range(ws.Cells(row + 1, sc), ws.Cells(row + 1, dc)).Merge
            ws.Cells(row + 1, sc).Formula2 = "=IFERROR(TEXTJOIN(CHAR(10),TRUE,UNIQUE(" & lst & "))," & Q(IIf(allMode, "No task", "No AoL task")) & ")"
            Fmt ws.Range(ws.Cells(row + 1, sc), ws.Cells(row + 1, dc)), False, 9, "", "", False, True, 0, True
            ws.Range(ws.Cells(row, sc), ws.Cells(row, dc)).Borders(xlEdgeTop).Weight = xlMedium
            ws.Range(ws.Cells(row, sc), ws.Cells(row, dc)).Borders(xlEdgeTop).Color = HexRGB("#6B7280")
            k = LineCount(CStr(years(y)), g, allMode)
            If k > maxLines Then maxLines = k
            ' status colours
            With ws.Cells(row, sc).FormatConditions
                If allMode Then
                    .Add(xlCellValue, xlEqual, "=0").Interior.Color = HexRGB("#F8D7D0")
                Else
                    .Add(xlCellValue, xlGreaterEqual, "=" & AOL_TARGET).Interior.Color = HexRGB("#D7EFD9")
                    .Add(xlCellValue, xlBetween, "=1", "=" & (AOL_TARGET - 1)).Interior.Color = HexRGB("#FDE7B0")
                    .Add(xlCellValue, xlEqual, "=0").Interior.Color = HexRGB("#F8D7D0")
                End If
            End With
        Next y
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 1)).Borders(xlEdgeTop).Weight = xlMedium
        ws.Rows(row).RowHeight = 38
        ws.Rows(row + 1).RowHeight = Application.Max(14, 12.5 * maxLines + 4)
        row = row + 2
    Next g
    ' hidden helper rows: unique-task counts per year by type code, then the short mix text
    Dim helpRow As Long, tot As String, parts As String
    helpRow = row + 2
    For y = LBound(years) To UBound(years)
        sc = 2 + 2 * (y - LBound(years))
        base = "(" & yearR & "=" & Q(CStr(years(y))) & ")*(" & visR & "=1)" & keep
        For k = 1 To TCount
            ws.Cells(helpRow + k - 1, sc).Formula2 = "=IFERROR(ROWS(UNIQUE(FILTER(" & lineR & "," & base & "*(TEXTAFTER(" & lineR & "," & Q(GSect()) & ",-1)=" & Q(TCode(k)) & ")))),0)"
        Next k
        ws.Cells(helpRow + TCount, sc).Formula2 = "=IFERROR(ROWS(UNIQUE(FILTER(" & lineR & "," & base & "))),0)"
        tot = ws.Cells(helpRow + TCount, sc).Address(False, False)
        parts = ""
        For k = 1 To TCount
            parts = parts & IIf(k > 1, "&", "") & "IF(" & ws.Cells(helpRow + k - 1, sc).Address(False, False) & ">0," & _
                Q(SepDot() & TCode(k) & " ") & "&TEXT(" & ws.Cells(helpRow + k - 1, sc).Address(False, False) & "/" & tot & ",""0%""),"""")"
        Next k
        ws.Range(ws.Cells(head + 1, sc), ws.Cells(head + 1, sc + 1)).Merge
        ws.Cells(head + 1, sc).Formula = "=IF(" & tot & "=0," & Q(GMdash()) & ",MID(" & parts & ",4,999))"
    Next y
    For k = 0 To TCount
        If k < TCount Then ws.Cells(helpRow + k, 1).Value = TCode(k + 1) Else ws.Cells(helpRow + k, 1).Value = "total"
    Next k
    ws.Range(ws.Cells(helpRow, 1), ws.Cells(helpRow + TCount, 1)).EntireRow.Hidden = True
    ' subtitle
    statusCells = Mid$(statusCells, 2)
    f = ""
    Dim cell As Variant
    For Each cell In Split(statusCells, ",")
        f = f & IIf(Len(f) > 0, "+", "") & "(" & cell & IIf(allMode, "=0)", ">=" & AOL_TARGET & ")")
    Next cell
    If allMode Then
        ws.Range("A2").Formula = "=""All tasks by GLO" & SepDot() & """&(" & f & ")&"" of " & (UBound(Split(statusCells, ",")) + 1) & _
            " GLO " & GTimes() & " year cells have no task" & SepDot() & "follows filters on the Tasks sheet"""
    Else
        ws.Range("A2").Formula = "=""Assessment balance" & SepDot() & """&(" & f & ")&"" of " & (UBound(Split(statusCells, ",")) + 1) & _
            " GLO " & GTimes() & " year cells have " & AOL_TARGET & "+ AoL tasks" & SepDot() & "follows filters on the Tasks sheet"""
    End If
    Fmt ws.Range("A2"), False, 0, "#6B7280", "", True
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(head + 2, 2).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    With ws.PageSetup
        .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
End Sub

Private Function LineCount(ByVal y As String, ByVal g As Long, ByVal allMode As Boolean) As Long
    Dim r As Long, seen As New Collection
    For r = 1 To NR
        If V(r, "Study Year") = y Then
            If allMode Then
                If Len(RAny(r, g)) > 0 Then AddUnique seen, RAll(r)
            Else
                If Len(RMarks(r, g)) > 0 And (RMark(r) = GFlag() Or RMark(r) = GDot()) Then AddUnique seen, RBal(r)
            End If
        End If
    Next r
    LineCount = seen.Count
End Function

' ========================================================================
'  SUMMARY (+ hidden Outline)
' ========================================================================
Private Sub BuildSummary()
    Dim wo As Worksheet, ws As Worksheet, keys() As Variant, sk() As Variant, i As Long, r As Long, n As Long
    Dim seen As New Collection, idx() As Long, out() As Variant, m As Long, gc As Variant, g As Long
    Dim curP As String, curU As String, rowsP As New Collection, typ() As String, rowOf() As Long
    gc = GloCols()
    ' one task row per key (year|tri|unit|AT), sorted
    For r = 1 To NR
        If Not HasKey(seen, "k" & RKey(r)) Then seen.Add r, "k" & RKey(r)
    Next r
    n = seen.Count
    ReDim keys(1 To n): ReDim sk(1 To n)
    For i = 1 To n
        r = seen(i): keys(i) = r
        sk(i) = YearSortKey(V(r, "Study Year")) & "|" & TriSortKey(V(r, "Trimester")) & "|" & V(r, "Unit Code") & "|" & NaturalPad(V(r, "AT Number"))
    Next i
    SortByKeys keys, sk
    ' outline rows: Period / Unit / Task
    ReDim typ(1 To n * 3): ReDim rowOf(1 To n * 3)
    m = 0
    For i = 1 To n
        r = keys(i)
        If V(r, "Study Year") & "|" & V(r, "Trimester") <> curP Then
            curP = V(r, "Study Year") & "|" & V(r, "Trimester"): curU = ""
            m = m + 1: typ(m) = "Period": rowOf(m) = r
        End If
        If V(r, "Unit Code") <> curU Then
            curU = V(r, "Unit Code")
            m = m + 1: typ(m) = "Unit": rowOf(m) = r
        End If
        m = m + 1: typ(m) = "Task": rowOf(m) = r
    Next i
    Set wo = NewSheet("Outline")
    ' A Show | B Type | C Year | D Tri | E Unit | F Task | G Title | H AoL | I Hurdle | J Weight
    ' K-R GLO codes | S Mode | T Group | U Tools | V AoL task type | W Key | X-AE ranks
    ReDim out(1 To m + 1, 1 To 31)
    Dim hdrs As Variant
    hdrs = Split("Show|Type|Year|Trimester|Unit|Task|Assessment title|AoL|Hurdle|Weight|G1|G2|G3|G4|G5|G6|G7|G8|Mode|Group|Tools|AoL task type|Key|R1|R2|R3|R4|R5|R6|R7|R8", "|")
    For i = 0 To 30: out(1, i + 1) = hdrs(i): Next i
    Dim code As String, rank As Long, w As Double
    For i = 1 To m
        r = rowOf(i)
        out(i + 1, 2) = typ(i)
        Select Case typ(i)
            Case "Period": out(i + 1, 3) = V(r, "Study Year"): out(i + 1, 4) = V(r, "Trimester")
            Case "Unit": out(i + 1, 3) = V(r, "Study Year"): out(i + 1, 4) = V(r, "Trimester"): out(i + 1, 5) = V(r, "Unit Code")
            Case "Task"
                out(i + 1, 6) = V(r, "AT Number"): out(i + 1, 7) = V(r, "Assessment Title")
                out(i + 1, 8) = IIf(LCase$(Trim$(V(r, "AoL"))) = "yes", "Yes", "No")
                out(i + 1, 9) = IIf(IsHurdleYes(V(r, "Hurdle")), "Yes", "No")
                w = ShortWeight(V(r, "Weighting"))
                If w >= 0 Then out(i + 1, 10) = w
                For g = 1 To 8
                    If gc(g) > 0 Then GloCode Dat(r, gc(g)), code, rank Else code = "": rank = 0
                    out(i + 1, 10 + g) = code: out(i + 1, 23 + g) = rank
                Next g
                out(i + 1, 19) = RMode(r): out(i + 1, 20) = RPart(r): out(i + 1, 21) = Trim$(V(r, "Tools")): out(i + 1, 22) = RAolShort(r)
                out(i + 1, 23) = RKey(r)
        End Select
    Next i
    wo.Range("C:W").NumberFormat = "@"
    wo.Range(wo.Cells(1, 1), wo.Cells(m + 1, 31)).Value = out
    wo.Range("C:W").NumberFormat = "General"       ' values stay text; formulas below must calculate
    wo.Range("J:J").NumberFormat = "0%"
    ' formulas
    Dim a As Long, b As Long, j As Long, showRng As String, cnt As String, labels As Variant, lab As Variant, colL As String
    For i = 1 To m
        r = i + 1
        If typ(i) = "Task" Then
            wo.Cells(r, 1).Formula = "=IF(COUNTIFS(" & TRange(tKey) & ",W" & r & "," & TRange(tVis) & ",1)>0,1,0)"
        Else
            j = i + 1
            Do While j <= m
                If typ(j) = "Period" Then Exit Do
                If typ(i) = "Unit" And typ(j) = "Unit" Then Exit Do
                j = j + 1
            Loop
            a = i + 2: b = j          ' Excel rows of first/last child
            showRng = "A" & a & ":A" & b
            wo.Cells(r, 1).Formula = "=IF(SUM(" & showRng & ")>0,1,0)"
            cnt = "COUNTIFS(B" & a & ":B" & b & ",""Task""," & showRng & ",1)"
            If typ(i) = "Period" Then
                wo.Cells(r, 5).Formula = "=COUNTIFS(B" & a & ":B" & b & ",""Unit""," & showRng & ",1)&"" units"""
                wo.Cells(r, 6).Formula = "=" & cnt & "&"" tasks"""
            Else
                wo.Cells(r, 6).Formula = "=" & cnt & "&IF(" & cnt & "=1,"" task"","" tasks"")"
                wo.Cells(r, 8).Formula = "=COUNTIFS(H" & a & ":H" & b & ",""Yes""," & showRng & ",1)"
                wo.Cells(r, 9).Formula = "=COUNTIFS(I" & a & ":I" & b & ",""Yes""," & showRng & ",1)"
                wo.Cells(r, 10).Formula = "=SUMIFS(J" & a & ":J" & b & "," & showRng & ",1)"
                For Each lab In Array("S", "T")
                    colL = CStr(lab)
                    If colL = "S" Then labels = Array("In person", "Online", "Both", "Other") Else labels = Array("Ind", "Grp", "Both")
                    Dim fx As String: fx = "=TEXTJOIN("" " & GMid() & " "",TRUE"
                    Dim lb As Variant
                    For Each lb In labels
                        fx = fx & ",IF(COUNTIFS(" & colL & a & ":" & colL & b & "," & Q(CStr(lb)) & "," & showRng & ",1)>0," & Q(CStr(lb) & " ") & "&COUNTIFS(" & colL & a & ":" & colL & b & "," & Q(CStr(lb)) & "," & showRng & ",1),"""")"
                    Next lb
                    wo.Range(colL & r).Formula2 = fx & ")"
                Next lab
                wo.Cells(r, 21).Formula2 = "=IFERROR(TEXTJOIN("", "",TRUE,UNIQUE(FILTER(U" & a & ":U" & b & ",(" & showRng & "=1)*(U" & a & ":U" & b & "<>"""")))),"""")"
            End If
            For g = 1 To 8
                colL = ColLetter(23 + g)
                wo.Cells(r, 10 + g).Formula = "=CHOOSE(MAXIFS(" & colL & a & ":" & colL & b & "," & showRng & ",1)+1,""""," & Q(GMid()) & ",""?"",""E"",""Y"",""H"")"
            Next g
        End If
    Next i
    ' Summary sheet
    Set ws = NewSheet("Summary")
    ws.Columns("A").ColumnWidth = 2: ws.Columns("A").Hidden = True
    ws.Columns("B").ColumnWidth = 9: ws.Columns("C").ColumnWidth = 12: ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 10: ws.Columns("F").ColumnWidth = 44: ws.Columns("G:H").ColumnWidth = 7
    ws.Columns("I").ColumnWidth = 7: ws.Columns("I").NumberFormat = "0%;;"
    ws.Columns("J:Q").ColumnWidth = 4.2
    ws.Columns("R").ColumnWidth = 13: ws.Columns("S").ColumnWidth = 11: ws.Columns("T").ColumnWidth = 14: ws.Columns("U").ColumnWidth = 32
    ws.Range("B1").Value = CourseTitle(): Fmt ws.Range("B1"), True, 14
    ws.Range("B2").Formula = "=""Showing ""&COUNTIFS(Outline!A2:A" & (m + 1) & ",1,Outline!B2:B" & (m + 1) & ",""Task"")&"" of " & n & " tasks" & SepDot() & "filter the Tasks sheet to change this view"""
    Fmt ws.Range("B2"), False, 0, "#6B7280", "", True
    ws.Range("B3").Value = "G1" & GNdash() & "G8: H assessed (hurdle) " & SepDot() & "Y assessed (non-hurdle) " & SepDot() & GMid() & " not assessed " & SepDot() & "E data error " & SepDot() & "? TBC. Unit rows roll up their visible tasks."
    Fmt ws.Range("B3"), False, 9, "#6B7280"
    ws.Range("A5:U5").Value = Array("", "Year", "Trimester", "Unit", "Task", "Assessment title", "AoL", "Hurdle", "Weight", _
        "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "Mode", "Group", "Tools", "AoL task type")
    Fmt ws.Range("A5:U5"), True
    ws.Range("A5:U5").Borders(xlEdgeBottom).Weight = xlThin
    ws.Range("J5:Q5").HorizontalAlignment = xlCenter
    ws.Range("A6").Formula2 = "=FILTER(IF(Outline!B2:V" & (m + 1) & "="""","""",Outline!B2:V" & (m + 1) & "),Outline!A2:A" & (m + 1) & "=1,""No tasks match the filters on the Tasks sheet"")"
    ws.Range(ws.Cells(6, 10), ws.Cells(6 + m, 17)).HorizontalAlignment = xlCenter
    Dim area As Range: Set area = ws.Range(ws.Cells(6, 1), ws.Cells(6 + m, 21))
    With ws.Range(ws.Cells(6, 10), ws.Cells(6 + m, 17)).FormatConditions
        .Add(xlCellValue, xlEqual, "=""H""").Interior.Color = HexRGB("#1F6F66")
        .Item(.Count).Font.Color = vbWhite
        .Add(xlCellValue, xlEqual, "=""Y""").Interior.Color = HexRGB("#BFE3DE")
        .Add(xlCellValue, xlEqual, "=""E""").Interior.Color = HexRGB("#E05252")
        .Add(xlCellValue, xlEqual, "=""?""").Interior.Color = HexRGB("#C9D3E0")
    End With
    With area.FormatConditions.Add(xlExpression, , "=$A6=""Period""")
        .Font.Bold = True: .Interior.Color = HexRGB("#E5E7EB")
    End With
    area.FormatConditions.Add(xlExpression, , "=$A6=""Unit""").Font.Bold = True
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
    ws.Range("B1").Select
End Sub

' ========================================================================
'  LEGEND
' ========================================================================
Private Sub BuildLegend()
    Dim ws As Worksheet, r As Long, mk As Variant, g As Long, k As Long
    Set ws = NewSheet("Legend")
    ws.Columns("A").ColumnWidth = 3: ws.Columns("B").ColumnWidth = 26: ws.Columns("C").ColumnWidth = 95
    ws.Range("B1").Value = "Quick guide": Fmt ws.Range("B1"), True, 14
    ws.Range("B2").Value = CourseTitle(): Fmt ws.Range("B2"), False, 0, "#6B7280", "", True
    r = 4
    LegendSection ws, r, "Buttons"
    LegendItem ws, r, "Import export...", "Load a dashboard timeline export (.xlsx or .csv) into Tasks. Planning columns already filled in are kept (matched by unit + AT)."
    LegendItem ws, r, "Refresh views", "Rebuild every sheet from the current Tasks data, after editing weeks, weightings, Tools, Assessment type or the Lists sheet."
    LegendSection ws, r, "Sheets"
    LegendItem ws, r, "Timeline", "Every task as a card by study year " & GTimes() & " trimester " & GTimes() & " AT1/AT2/AT3/Other (not due dates)."
    LegendItem ws, r, "Weeks", "The same cards by week: Planned Weeks first, then unit-guide weeks, then the AT number (dashed, not a due date)."
    LegendItem ws, r, "AOL type balance", "GLO " & GTimes() & " year level. Target: every GLO in at least 2 AoL tasks (" & GFlag() & " or " & GDot() & ") at every level."
    LegendItem ws, r, "All tasks", "Same layout, every task under each GLO it addresses (unit-guide alignment where available, otherwise CALM)."
    LegendItem ws, r, "Summary", "Year/trimester " & GArrow() & " unit " & GArrow() & " task list with GLO codes, mode, group, tools and AoL task type."
    LegendItem ws, r, "Tasks", "The dashboard export, one row per task. Filter here and every view follows. Planning columns are typed in here."
    LegendItem ws, r, "Lists", "Task type codes and their matching words. Edit, then Refresh views."
    LegendSection ws, r, "Alignment marks"
    For Each mk In AllMarks()
        ws.Cells(r, 2).Value = mk & "  " & MarkLabel(CStr(mk))
        Fmt ws.Cells(r, 2), True, 10, "", MarkFill(CStr(mk))
        ws.Cells(r, 3).Value = Choose(Application.Match(mk, AllMarks(), 0), _
            "AoL task that is also a hurdle. Counts towards the AoL target.", "AoL task, not a hurdle. Counts towards the AoL target.", _
            "Hurdle task, not AoL. Does not count towards the AoL target.", "Neither AoL nor hurdle.")
        r = r + 1
    Next mk
    LegendSection ws, r, "GLO colours"
    For g = 1 To 8
        ws.Cells(r, 2).Value = GDot() & " GLO" & g: Fmt ws.Cells(r, 2), True, 10, GloColour(g)
        ws.Cells(r, 3).Value = GloName(g)
        r = r + 1
    Next g
    LegendSection ws, r, "Task type codes (Lists sheet)"
    For k = 1 To TCount
        LegendItem ws, r, TCode(k), TLabel(k) & IIf(k < TCount, "  (matches: " & TWords(k) & ")", "  (none of the above)")
    Next k
    LegendSection ws, r, "Filters (Timeline and Weeks)"
    LegendItem ws, r, "Click a GLO", "Toggles that GLO (" & GBoxOn() & "/" & GBoxOff() & "). Ticked GLOs show tasks aligned to any of them; untick all to show everything."
    LegendItem ws, r, "Click a mode", "In person, Online, Other / unspecified (any of the ticked)."
    LegendItem ws, r, "Weighting from" & GNdash() & "to", "Type a % range; tasks with unreadable weights only show at 0" & GNdash() & "100."
    LegendItem ws, r, "Tasks filters", "Filter any Tasks column (e.g. a pathway's " & GDot() & " column for one major). Every sheet follows."
    LegendSection ws, r, "Planning columns (Tasks sheet)"
    LegendItem ws, r, "Tools", "Tech used, e.g. FFIO, OT. Shown on cards and in the Summary."
    LegendItem ws, r, "Assessment type", "The new task definitions. Used for type codes and balance lines instead of CALM."
    LegendItem ws, r, "Planned Weeks", "Planned submission week(s), e.g. 6 or 6; 10. Takes priority on the Weeks sheet."
End Sub

Private Sub LegendSection(ByVal ws As Worksheet, ByRef r As Long, ByVal title As String)
    r = r + 1
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Merge
    ws.Cells(r, 2).Value = title
    Fmt ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)), True, 11, "#303842"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Borders(xlEdgeTop).Weight = xlMedium
    r = r + 1
End Sub

Private Sub LegendItem(ByVal ws As Worksheet, ByRef r As Long, ByVal k As String, ByVal v As String)
    ws.Cells(r, 2).Value = k: ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = v: ws.Cells(r, 3).WrapText = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).VerticalAlignment = xlTop
    r = r + 1
End Sub

' ---- buttons -------------------------------------------------------------
Private Sub AddButtons(ByVal ws As Worksheet, ByVal leftPos As Double, ByVal topPos As Double, ByVal onGrid As Boolean)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, 110, 22)
    s.Name = "btnImport": s.OnAction = "ImportExport"
    s.TextFrame2.TextRange.Text = "Import export" & GEll()
    StyleButton s, "#2B2F3A"
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos + 118, topPos, 100, 22)
    s.Name = "btnRefresh": s.OnAction = "RefreshViewsButton"
    s.TextFrame2.TextRange.Text = "Refresh views"
    StyleButton s, "#1F6F66"
End Sub

Private Sub StyleButton(ByVal s As Shape, ByVal fillHex As String)
    s.Fill.ForeColor.RGB = HexRGB(fillHex)
    s.Line.Visible = msoFalse
    With s.TextFrame2.TextRange.Font
        .Size = 10: .Bold = msoTrue: .Fill.ForeColor.RGB = vbWhite
    End With
    s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    s.TextFrame2.VerticalAnchor = msoAnchorMiddle
    s.Placement = xlFreeFloating
End Sub
