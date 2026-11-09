Attribute VB_Name = "modData"
' ========================================================================
'  Curriculum Map Builder - data: import, Tasks table, per-row derived fields
' ========================================================================
Option Explicit

Public Const TASKS_SHEET As String = "Tasks"
Public Const PLANNING_LIST As String = "Tools|Assessment type|Planned Weeks"
Public Const FALLBACK_LINE_TEXT As String = " Week from AT number"   ' prefixed with the approx sign

' ---- the loaded table (module state used by the view builders) ----------
Public Hdr() As String          ' data column headers (1..NC)
Public Dat() As String          ' data values (1..NR, 1..NC)
Public NR As Long, NC As Long

' derived per row (1..NR)
Public RMark() As String, RCard() As String, RBal() As String, RAll() As String, RSort() As String
Public RKey() As String, RMarks() As String, RAny() As String      ' RMarks/RAny: (row, 1..8)
Public RIp() As Long, ROn() As Long, ROth() As Long, RWt() As Variant
Public RMode() As String, RPart() As String, RAolShort() As String
Public RHasCalmGlo() As Boolean, RHasAnyGlo() As Boolean, RHasGuideGlo() As Boolean

' type codes (from the Lists sheet)
Public TCode() As String, TLabel() As String, TWords() As String, TCount As Long

' ---- column lookup -------------------------------------------------------
Public Function ColOf(ParamArray names() As Variant) As Long
    Dim i As Long, j As Long
    For j = LBound(names) To UBound(names)
        For i = 1 To NC
            If LCase$(Hdr(i)) = LCase$(names(j)) Then ColOf = i: Exit Function
        Next i
    Next j
    For j = LBound(names) To UBound(names)            ' prefix match ("Assessment Title (...)")
        For i = 1 To NC
            If LCase$(Left$(Hdr(i), Len(names(j)))) = LCase$(names(j)) Then ColOf = i: Exit Function
        Next i
    Next j
End Function

Public Function V(ByVal r As Long, ByVal colName As String) As String
    Dim c As Long: c = ColOf(colName)
    If c > 0 Then V = Dat(r, c)
End Function

Public Function GloCols() As Variant
    ' the 8 export columns "GLO1 (...)" .. "GLO8 (...)"
    Dim out(1 To 8) As Long, i As Long, g As Long
    For i = 1 To NC
        If Left$(Hdr(i), 3) = "GLO" And Len(Hdr(i)) > 4 Then
            If IsDigitChar(Mid$(Hdr(i), 4, 1)) And Mid$(Hdr(i), 5, 1) <> "M" And Mid$(Hdr(i), 5, 2) <> " M" _
               And Mid$(Hdr(i), 5, 2) <> " A" Then
                g = CLng(Mid$(Hdr(i), 4, 1))
                If g >= 1 And g <= 8 And out(g) = 0 Then out(g) = i
            End If
        End If
    Next i
    GloCols = out
End Function

Public Function HelperNames() As Variant
    Dim a As String, g As Long
    For g = 1 To 8: a = a & "GLO" & g & " Mark|": Next g
    For g = 1 To 8: a = a & "GLO" & g & " Any|": Next g
    a = a & "In person|Online|Other mode|Weight %|Balance Line|All Line|Sort Key|Row ID|Card|Map Key|Visible|Show (Timeline)"
    HelperNames = Split(a, "|")
End Function

Private Function IsHelper(ByVal h As String) As Boolean
    Dim x As Variant
    For Each x In HelperNames()
        If h = x Then IsHelper = True: Exit Function
    Next x
End Function

' ---- pathways -> flag column labels --------------------------------------
Public Function PathwayList() As Collection
    Dim out As New Collection, r As Long, p As Variant, c As Long, s As String
    c = ColOf("Pathways")
    If c = 0 Then Set PathwayList = out: Exit Function
    For r = 1 To NR
        For Each p In Split(Dat(r, c), ";")
            s = Trim$(p)
            If Len(s) > 0 Then AddUnique out, s
        Next p
    Next r
    Set PathwayList = out
End Function

Public Function CourseCount() As Long
    Dim c As Long, r As Long, seen As New Collection
    c = ColOf("Course Code")
    If c = 0 Then CourseCount = 1: Exit Function
    For r = 1 To NR: AddUnique seen, Dat(r, c): Next r
    CourseCount = seen.Count
End Function

Public Function PathHeader(ByVal p As String, ByVal oneCourse As Boolean) As String
    Dim k As Long, s As String
    s = p
    If oneCourse Then
        k = InStr(s, "_")
        If k > 0 Then s = Mid$(s, k + 1)
    End If
    PathHeader = Left$(Replace(s, " : ", " "), 250)
End Function

' ========================================================================
'  IMPORT
' ========================================================================
Public Sub ImportExport()
    Dim f As Variant
    f = Application.GetOpenFilename()
    If VarType(f) = vbBoolean Then Exit Sub
    ImportExportFrom CStr(f)
End Sub

' Load a dashboard timeline export (.csv or .xlsx) into Tasks, keeping planning
' and team-added columns already filled in (matched by year|trimester|unit|AT,
' then unit|AT), then rebuild every view.
Public Function ImportExportFrom(ByVal path As String) As String
    Dim newHdr() As String, newDat() As String, nr2 As Long, nc2 As Long
    Dim oldHdr() As String, oldDat() As String, oldNR As Long, oldNC As Long, haveOld As Boolean
    On Error GoTo Fail
    If Not ReadExportFile(path, newHdr, newDat, nr2, nc2) Then
        ImportExportFrom = "ERR could not read " & path: Exit Function
    End If
    ' previous data (planning columns etc.)
    haveOld = ReadTasksData(oldHdr, oldDat, oldNR, oldNC)
    ' columns to carry over: in the old table but not in the new export, plus planning columns
    Dim extra As New Collection, h As Variant, i As Long, j As Long
    If haveOld Then
        For i = 1 To oldNC
            If Not InArr(newHdr, nc2, oldHdr(i)) And Not IsExportColumn(oldHdr(i)) Then AddUnique extra, oldHdr(i)
        Next i
    End If
    For Each h In Split(PLANNING_LIST, "|")
        If Not InArr(newHdr, nc2, CStr(h)) Then AddUnique extra, CStr(h)
    Next h
    ' build the combined table
    NC = nc2 + extra.Count: NR = nr2
    ReDim Hdr(1 To NC): ReDim Dat(1 To IIf(NR > 0, NR, 1), 1 To NC)
    For i = 1 To nc2: Hdr(i) = newHdr(i): Next i
    For i = 1 To extra.Count: Hdr(nc2 + i) = extra(i): Next i
    For i = 1 To NR
        For j = 1 To nc2: Dat(i, j) = newDat(i, j): Next j
    Next i
    If haveOld And extra.Count > 0 Then CarryOver oldHdr, oldDat, oldNR, oldNC, nc2, extra
    ' planning columns that exist in the new export but are blank there: keep old values
    If haveOld Then FillBlankPlanning oldHdr, oldDat, oldNR, oldNC
    WriteTasksData
    ImportExportFrom = RefreshViews()
    Exit Function
Fail:
    ImportExportFrom = "ERR " & Err.Number & " " & Err.Description
End Function

' columns the dashboard export itself produces: never carried over from an older import
Private Function IsExportColumn(ByVal h As String) As Boolean
    Dim x As Variant
    For Each x In Split("Study Year|Trimester|AT Group|Unit Code|AT Number|Task Number|Course Code|Course|Pathways|" & _
                        "Teaching Period Source|AoL Task Type|AoL|Hurdle|Weighting|Mode|Participation|Student Output|" & _
                        "Task Alignment|Alignment Mark|Aligned GLOs|Guide Weeks|Guide Week Detail|Guide File|" & _
                        "Guide Evidence|Guide ULOs|Guide GLOs", "|")
        If LCase$(h) = LCase$(x) Then IsExportColumn = True: Exit Function
    Next x
    If LCase$(Left$(h, 16)) = "assessment title" Then IsExportColumn = True: Exit Function
    If Left$(h, 3) = "GLO" And InStr(h, "(") > 0 Then IsExportColumn = True
End Function

Private Function InArr(a() As String, n As Long, ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To n
        If LCase$(a(i)) = LCase$(s) Then InArr = True: Exit Function
    Next i
End Function

Private Function RowKey(h() As String, d() As String, nc_ As Long, r As Long, full As Boolean) As String
    Dim cy As Long, ct As Long, cu As Long, ca As Long, i As Long
    For i = 1 To nc_
        Select Case LCase$(h(i))
            Case "study year": cy = i
            Case "trimester": ct = i
            Case "unit code": cu = i
            Case "at number", "task number": If ca = 0 Then ca = i
        End Select
    Next i
    If cu = 0 Or ca = 0 Then Exit Function
    If full And cy > 0 And ct > 0 Then
        RowKey = UCase$(d(r, cy) & "|" & d(r, ct) & "|" & d(r, cu) & "|" & d(r, ca))
    Else
        RowKey = UCase$(d(r, cu) & "|" & d(r, ca))
    End If
End Function

Private Sub CarryOver(oh() As String, od() As String, onr As Long, onc As Long, nc2 As Long, extra As Collection)
    Dim byFull As New Collection, byTask As New Collection, r As Long, k As String, i As Long, oc As Long, src As Long, k2 As Long
    For r = 1 To onr
        k = RowKey(oh, od, onc, r, True)
        If Len(k) > 0 And Not HasKey(byFull, "k" & k) Then byFull.Add r, "k" & k
        k = RowKey(oh, od, onc, r, False)
        If Len(k) > 0 And Not HasKey(byTask, "k" & k) Then byTask.Add r, "k" & k
    Next r
    For r = 1 To NR
        src = 0
        k = RowKey(Hdr, Dat, NC, r, True)
        If HasKey(byFull, "k" & k) Then src = byFull("k" & k)
        If src = 0 Then
            k = RowKey(Hdr, Dat, NC, r, False)
            If HasKey(byTask, "k" & k) Then src = byTask("k" & k)
        End If
        If src > 0 Then
            For i = 1 To extra.Count
                oc = 0
                For k2 = 1 To onc
                    If oh(k2) = extra(i) Then oc = k2: Exit For
                Next k2
                If oc > 0 Then Dat(r, nc2 + i) = od(src, oc)
            Next i
        End If
    Next r
End Sub

Private Sub FillBlankPlanning(oh() As String, od() As String, onr As Long, onc As Long)
    Dim h As Variant, c As Long, oc As Long, r As Long, src As Long, k As String, i As Long
    Dim byTask As New Collection
    For r = 1 To onr
        k = RowKey(oh, od, onc, r, False)
        If Len(k) > 0 And Not HasKey(byTask, "k" & k) Then byTask.Add r, "k" & k
    Next r
    For Each h In Split(PLANNING_LIST, "|")
        c = ColOf(CStr(h)): oc = 0
        For i = 1 To onc
            If LCase$(oh(i)) = LCase$(h) Then oc = i: Exit For
        Next i
        If c > 0 And oc > 0 Then
            For r = 1 To NR
                If Len(Dat(r, c)) = 0 Then
                    k = RowKey(Hdr, Dat, NC, r, False)
                    If HasKey(byTask, "k" & k) Then
                        src = byTask("k" & k)
                        Dat(r, c) = od(src, oc)
                    End If
                End If
            Next r
        End If
    Next h
End Sub

Private Function ReadExportFile(ByVal path As String, h() As String, d() As String, nr_ As Long, nc_ As Long) As Boolean
    Dim v As Variant, r As Long, c As Long
    If LCase$(Right$(path, 4)) = ".csv" Then
        v = ParseCsv(ReadUtf8File(path))
        If IsEmpty(v) Then Exit Function
    Else
        Dim wb As Workbook, ws As Worksheet, s As Worksheet
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
        For Each s In wb.Worksheets
            If LCase$(s.Name) = "export" Or LCase$(s.Name) = "tasks" Then Set ws = s: Exit For
        Next s
        If ws Is Nothing Then Set ws = wb.Worksheets(1)
        v = ws.UsedRange.Value
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    ' trim trailing empty rows
    nc_ = UBound(v, 2)
    Do While nc_ > 1 And Len(Trim$(CStr(v(1, nc_)))) = 0: nc_ = nc_ - 1: Loop
    nr_ = UBound(v, 1) - 1
    Do While nr_ > 0
        If Len(Trim$(CStr(v(nr_ + 1, 1)))) > 0 Then Exit Do
        nr_ = nr_ - 1
    Loop
    ReDim h(1 To nc_): ReDim d(1 To IIf(nr_ > 0, nr_, 1), 1 To nc_)
    For c = 1 To nc_: h(c) = Trim$(CStr(v(1, c))): Next c
    For r = 1 To nr_
        For c = 1 To nc_: d(r, c) = CStr(v(r + 1, c)): Next c
    Next r
    ' the export must look like a dashboard timeline export
    ReadExportFile = InArr(h, nc_, "Unit Code") And (InArr(h, nc_, "AT Number") Or InArr(h, nc_, "Task Number"))
End Function

' Read the data columns of the current Tasks table (skipping helper and pathway columns)
Public Function ReadTasksData(h() As String, d() As String, nr_ As Long, nc_ As Long) As Boolean
    Dim ws As Worksheet, lo As ListObject, v As Variant, r As Long, c As Long, keep() As Long, k As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TASKS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        If lo.DataBodyRange Is Nothing Then Exit Function
        v = lo.Range.Value
    Else
        If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Function
        v = ws.UsedRange.Value
    End If
    ' pathway labels to skip (from the Pathways column)
    Dim pathCol As Long, flags As New Collection, p As Variant, oneCourse As Boolean, ccCol As Long, cc As New Collection
    For c = 1 To UBound(v, 2)
        If CStr(v(1, c)) = "Pathways" Then pathCol = c
        If CStr(v(1, c)) = "Course Code" Then ccCol = c
    Next c
    If ccCol > 0 Then
        For r = 2 To UBound(v, 1): AddUnique cc, CStr(v(r, ccCol)): Next r
    End If
    oneCourse = (cc.Count <= 1)
    If pathCol > 0 Then
        For r = 2 To UBound(v, 1)
            For Each p In Split(CStr(v(r, pathCol)), ";")
                If Len(Trim$(p)) > 0 Then AddUnique flags, Trim$(PathHeader(Trim$(p), oneCourse))
            Next p
        Next r
    End If
    ReDim keep(1 To UBound(v, 2))
    For c = 1 To UBound(v, 2)
        If Not IsHelper(CStr(v(1, c))) And Not HasKey(flags, "k" & Trim$(CStr(v(1, c)))) Then
            k = k + 1: keep(k) = c
        End If
    Next c
    nc_ = k: nr_ = UBound(v, 1) - 1
    ReDim h(1 To nc_): ReDim d(1 To IIf(nr_ > 0, nr_, 1), 1 To nc_)
    For c = 1 To nc_: h(c) = CStr(v(1, keep(c))): Next c
    For r = 1 To nr_
        For c = 1 To nc_: d(r, c) = CStr(v(r + 1, keep(c))): Next c
    Next r
    ReadTasksData = (nr_ > 0)
End Function

' Write Hdr/Dat as the Tasks sheet's data (helpers are added by RefreshViews)
Public Sub WriteTasksData()
    Dim ws As Worksheet, out() As Variant, r As Long, c As Long
    Set ws = EnsureSheet(TASKS_SHEET)
    ws.Cells.Clear
    Do While ws.ListObjects.Count > 0: ws.ListObjects(1).Delete: Loop
    ReDim out(1 To NR + 1, 1 To NC)
    For c = 1 To NC: out(1, c) = Hdr(c): Next c
    For r = 1 To NR
        For c = 1 To NC: out(r + 1, c) = Dat(r, c): Next c
    Next r
    ws.Range(ws.Cells(1, 1), ws.Cells(NR + 1, NC)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(NR + 1, NC)).Value = out
End Sub

Public Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function

' ========================================================================
'  TYPE CODES (editable Lists sheet)
' ========================================================================
Public Sub LoadTypeCodes()
    Dim ws As Worksheet, r As Long
    Set ws = EnsureSheet("Lists")
    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ws.Range("A1:C1").Value = Array("Code", "Label", "Matching words (comma-separated; ^ = must start a word)")
        ws.Range("A2:C2").Value = Array("W", "Written", "written, report, essay, exam, test, quiz, portfolio, assignment, reflect, paper, proposal, journal")
        ws.Range("A3:C3").Value = Array("IO", "Interactive oral", "interactive oral, viva, interview, ^oral")
        ws.Range("A4:C4").Value = Array("Pres", "Presentation", "presentation, ^present, poster, pitch, seminar, podcast, video")
        ws.Range("A5:C5").Value = Array("QA", "Q&A", "q&a, q & a, question and answer, questions and answers, questioning")
        ws.Range("A6:C6").Value = Array("Prac", "Practical", "practical, ^lab, skill, demonstrat, simulation, placement, field, experiment, clinical")
        ws.Range("A7:C7").Value = Array("Perf", "Performance", "performance, ^perform, exhibition, showcase, recital")
        ws.Range("A8:C8").Value = Array("Oth", "Other", "(anything else - keep this row last)")
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 8: ws.Columns("B").ColumnWidth = 18: ws.Columns("C").ColumnWidth = 90
        ws.Range("A10").Value = "Add a type: insert a row above 'Oth' with a short code, a label and the words to match, then click Refresh views."
        ws.Range("A10").Font.Italic = True
    End If
    TCount = 0
    ReDim TCode(1 To 50): ReDim TLabel(1 To 50): ReDim TWords(1 To 50)
    r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 And r < 52
        TCount = TCount + 1
        TCode(TCount) = Trim$(CStr(ws.Cells(r, 1).Value))
        TLabel(TCount) = Trim$(CStr(ws.Cells(r, 2).Value))
        TWords(TCount) = LCase$(CStr(ws.Cells(r, 3).Value))
        r = r + 1
    Loop
End Sub

' the type whose words appear FIRST in the text; the last row is the fallback
Public Function TypeCodeOf(ByVal text As String) As String
    Dim t As String, i As Long, w As Variant, word As String, p As Long, best As Long, bestCode As String
    t = LCase$(text)
    best = 0
    For i = 1 To TCount - 1
        For Each w In Split(TWords(i), ",")
            word = Trim$(w)
            If Len(word) > 0 Then
                p = FindWord(t, word)
                If p > 0 And (best = 0 Or p < best) Then best = p: bestCode = TCode(i)
            End If
        Next w
    Next i
    If best = 0 Then bestCode = TCode(TCount)
    TypeCodeOf = bestCode
End Function

Private Function FindWord(ByVal t As String, ByVal word As String) As Long
    Dim wordStart As Boolean, p As Long
    If Left$(word, 1) = "^" Then wordStart = True: word = Mid$(word, 2)
    p = InStr(t, word)
    Do While p > 0
        If Not wordStart Then FindWord = p: Exit Function
        If p = 1 Then FindWord = p: Exit Function
        If Not IsLetterChar(Mid$(t, p - 1, 1)) Then FindWord = p: Exit Function
        p = InStr(p + 1, t, word)
    Loop
End Function

' ========================================================================
'  DERIVED FIELDS
' ========================================================================
Public Sub DeriveRows()
    Dim r As Long, g As Long, gc As Variant, code As String, rank As Long
    Dim cU As Long, cA As Long, cT As Long, cW As Long, cMode As Long, multi As Boolean
    Dim ids As Collection, anyIds As Collection, x As Variant, title As String, w As Double, bits As String
    Dim lines As String, typeText As String, so As String, kind As String, wtxt As String
    gc = GloCols()
    cU = ColOf("Unit Code"): cA = ColOf("AT Number", "Task Number")
    cT = ColOf("Assessment Title"): cW = ColOf("Weighting"): cMode = ColOf("Mode")
    multi = (CourseCount() > 1)
    ReDim RMark(1 To NR): ReDim RCard(1 To NR): ReDim RBal(1 To NR): ReDim RAll(1 To NR): ReDim RSort(1 To NR)
    ReDim RKey(1 To NR): ReDim RMarks(1 To NR, 1 To 8): ReDim RAny(1 To NR, 1 To 8)
    ReDim RIp(1 To NR): ReDim ROn(1 To NR): ReDim ROth(1 To NR): ReDim RWt(1 To NR)
    ReDim RMode(1 To NR): ReDim RPart(1 To NR): ReDim RAolShort(1 To NR)
    ReDim RHasCalmGlo(1 To NR): ReDim RHasAnyGlo(1 To NR): ReDim RHasGuideGlo(1 To NR)
    For r = 1 To NR
        RMark(r) = MarkOf(V(r, "Alignment Mark"), V(r, "AoL"), V(r, "Hurdle"))
        ' CALM GLOs: "Aligned GLOs" (v7.2+) or the GLO1..8 columns with a "Yes" state
        Set ids = New Collection
        If ColOf("Aligned GLOs") > 0 And Len(Trim$(V(r, "Aligned GLOs") & V(r, "Alignment Mark"))) > 0 Then
            For Each x In Numbers(V(r, "Aligned GLOs"))
                If x >= 1 And x <= 8 Then AddUnique ids, CStr(x)
            Next x
        Else
            For g = 1 To 8
                If gc(g) > 0 Then
                    GloCode Dat(r, gc(g)), code, rank
                    If code = "H" Or code = "Y" Then AddUnique ids, CStr(g)
                End If
            Next g
        End If
        Set anyIds = New Collection
        For Each x In Numbers(V(r, "Guide GLOs"))
            If x >= 1 And x <= 8 Then AddUnique anyIds, CStr(x)
        Next x
        RHasGuideGlo(r) = (anyIds.Count > 0)
        If anyIds.Count = 0 Then Set anyIds = ids
        RHasCalmGlo(r) = (ids.Count > 0)
        RHasAnyGlo(r) = (anyIds.Count > 0)
        For g = 1 To 8
            If HasKey(ids, "k" & g) Then RMarks(r, g) = RMark(r) Else RMarks(r, g) = ""
            If HasKey(anyIds, "k" & g) Then RAny(r, g) = RMark(r) Else RAny(r, g) = ""
        Next g
        ' filters
        RMode(r) = ModeLabel(V(r, "Mode"))
        RIp(r) = IIf(RMode(r) = "In person" Or RMode(r) = "Both", 1, 0)
        ROn(r) = IIf(RMode(r) = "Online" Or RMode(r) = "Both", 1, 0)
        ROth(r) = IIf(RMode(r) = "Other", 1, 0)
        w = ShortWeight(V(r, "Weighting"))
        If w >= 0 Then RWt(r) = Round(w * 100, 2): wtxt = Format$(w * 100, "0.##") & "%" Else RWt(r) = "": wtxt = GMdash()
        If Right$(wtxt, 2) = ".%" Then wtxt = Left$(wtxt, Len(wtxt) - 2) & "%"
        RPart(r) = ParticipationLabel(V(r, "Participation"))
        RAolShort(r) = AolTypeShort(V(r, "AoL Task Type"))
        ' balance / all lines:  "<mark> UNIT . AT . type . weight[ . CALM] <sect><code>"
        typeText = Trim$(V(r, "Assessment type"))
        If Len(typeText) > 0 Then
            kind = typeText
        Else
            so = Squash(V(r, "Student Output"))
            If Len(so) = 0 Then so = GMdash()
            kind = Clip(so, 26) & "*"
            typeText = V(r, "Student Output")
        End If
        code = TypeCodeOf(typeText)
        RBal(r) = RMark(r) & " " & Dat(r, cU) & SepDot() & Dat(r, cA) & SepDot() & kind & SepDot() & wtxt
        RAll(r) = RBal(r) & IIf(RHasGuideGlo(r), "", SepDot() & "CALM") & " " & GSect() & code
        RBal(r) = RBal(r) & " " & GSect() & code
        RSort(r) = NaturalPad(Dat(r, cU)) & "|" & NaturalPad(Dat(r, cA))
        RKey(r) = V(r, "Study Year") & "|" & V(r, "Trimester") & "|" & Dat(r, cU) & "|" & Dat(r, cA)
        RCard(r) = CardText(r, "", multi, ids.Count > 0, wtxt)
    Next r
End Sub

Public Function CardText(ByVal r As Long, ByVal weekLine As String, ByVal multi As Boolean, _
                         ByVal hasCalmGlo As Boolean, ByVal wtxt As String) As String
    Dim lines As String, bits As String, title As String, mode As String
    title = Trim$(V(r, "Assessment Title"))
    If Len(title) = 0 Then title = "(no title)"
    mode = Clip(V(r, "Mode"), 22)
    If Len(mode) = 0 Then mode = GMdash()
    bits = wtxt & SepDot() & mode
    If multi Then bits = bits & SepDot() & V(r, "Course Code")
    lines = RMark(r) & " " & V(r, "Unit Code") & SepDot() & V(r, "AT Number") & vbLf & Clip(title, 46) & vbLf & bits
    If Not hasCalmGlo Then lines = lines & vbLf & "No GLO alignment"
    If Len(Trim$(V(r, "Assessment type"))) > 0 Then lines = lines & vbLf & "Type: " & Trim$(V(r, "Assessment type"))
    If Len(Trim$(V(r, "Tools"))) > 0 Then lines = lines & vbLf & "Tools: " & Trim$(V(r, "Tools"))
    If Len(weekLine) = 0 And Len(V(r, "Guide Week Detail")) > 0 Then weekLine = "Guide: " & V(r, "Guide Week Detail")
    If Len(weekLine) > 0 Then lines = lines & vbLf & weekLine
    CardText = lines
End Function

Public Function FallbackLine() As String
    FallbackLine = GApprox() & FALLBACK_LINE_TEXT
End Function
