Attribute VB_Name = "modUtil"
' ========================================================================
'  Curriculum Map Builder - utilities
'  Pure ASCII source: every non-ASCII character is built with ChrW, because
'  Mac Excel imports .bas files in a legacy encoding.
'  No Scripting.Dictionary / VBScript.RegExp (not available on Mac).
' ========================================================================
Option Explicit

' ---- glyphs --------------------------------------------------------------
Public Function GFlag() As String: GFlag = ChrW(&H2691): End Function      ' AoL + hurdle
Public Function GDot() As String: GDot = ChrW(&H25CF): End Function        ' AoL
Public Function GDiamond() As String: GDiamond = ChrW(&H25C6): End Function ' hurdle only
Public Function GRing() As String: GRing = ChrW(&H25CB): End Function      ' neither
Public Function GMid() As String: GMid = ChrW(&HB7): End Function           ' middle dot
Public Function GSect() As String: GSect = ChrW(&HA7): End Function         ' section sign
Public Function GEll() As String: GEll = ChrW(&H2026): End Function         ' ellipsis
Public Function GApprox() As String: GApprox = ChrW(&H2248): End Function
Public Function GArrow() As String: GArrow = ChrW(&H2192): End Function
Public Function GNdash() As String: GNdash = ChrW(&H2013): End Function
Public Function GMdash() As String: GMdash = ChrW(&H2014): End Function
Public Function GBoxOn() As String: GBoxOn = ChrW(&H2611): End Function
Public Function GBoxOff() As String: GBoxOff = ChrW(&H2610): End Function
Public Function GTimes() As String: GTimes = ChrW(&HD7): End Function
Public Function SepDot() As String: SepDot = " " & GMid() & " ": End Function

' ---- colours -------------------------------------------------------------
Public Function HexRGB(ByVal hexColour As String) As Long
    hexColour = Replace(hexColour, "#", "")
    HexRGB = RGB(CLng("&H" & Mid$(hexColour, 1, 2)), CLng("&H" & Mid$(hexColour, 3, 2)), CLng("&H" & Mid$(hexColour, 5, 2)))
End Function

Public Function GloColour(ByVal g As Long) As String
    GloColour = Choose(g, "#E76F51", "#F4A261", "#E9C46A", "#2A9D8F", "#607D8B", "#A06CD5", "#6D9DC5", "#C77DBA")
End Function

Public Function GloName(ByVal g As Long) As String
    GloName = Choose(g, "Discipline-specific knowledge", "Communication", "Digital literacy", "Critical thinking", _
                     "Problem solving", "Self-management", "Teamwork", "Global citizenship")
End Function

Public Function MarkFill(ByVal mark As String) As String
    Select Case mark
        Case GFlag(): MarkFill = "#F8D7D0"
        Case GDot(): MarkFill = "#E4DCF3"
        Case GDiamond(): MarkFill = "#F5E3C6"
        Case Else: MarkFill = "#F1F3F5"
    End Select
End Function

Public Function MarkLabel(ByVal mark As String) As String
    Select Case mark
        Case GFlag(): MarkLabel = "AoL + hurdle"
        Case GDot(): MarkLabel = "AoL, non-hurdle"
        Case GDiamond(): MarkLabel = "Hurdle, non-AoL"
        Case Else: MarkLabel = "Neither"
    End Select
End Function

Public Function AllMarks() As Variant
    AllMarks = Array(GFlag(), GDot(), GDiamond(), GRing())
End Function

' ---- columns -------------------------------------------------------------
Public Function ColLetter(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

' absolute "Sheet!$C$2:$C$n" for column c, rows 2..lastRow
Public Function ColRange(ByVal sheetName As String, ByVal c As Long, ByVal lastRow As Long) As String
    ColRange = QuoteSheet(sheetName) & "!$" & ColLetter(c) & "$2:$" & ColLetter(c) & "$" & lastRow
End Function

Public Function QuoteSheet(ByVal sheetName As String) As String
    If InStr(sheetName, " ") > 0 Then QuoteSheet = "'" & sheetName & "'" Else QuoteSheet = sheetName
End Function

' ---- text helpers --------------------------------------------------------
Public Function Clip(ByVal s As String, ByVal n As Long) As String
    s = Trim$(s)
    If Len(s) <= n Then Clip = s Else Clip = RTrim$(Left$(s, n - 1)) & GEll()
End Function

Public Function Squash(ByVal s As String) As String
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    Squash = Trim$(s)
End Function

Public Function IsDigitChar(ByVal ch As String) As Boolean
    IsDigitChar = (Len(ch) = 1 And ch >= "0" And ch <= "9")
End Function

Public Function IsLetterChar(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    IsLetterChar = (LCase$(ch) <> UCase$(ch))
End Function

' all integers in text, in order (e.g. "Wk 7 (Lab A), Wk 12" -> 7, 12)
Public Function Numbers(ByVal s As String) As Collection
    Dim out As New Collection, i As Long, cur As String, ch As String
    For i = 1 To Len(s) + 1
        If i <= Len(s) Then ch = Mid$(s, i, 1) Else ch = " "
        If IsDigitChar(ch) Then
            cur = cur & ch
        ElseIf Len(cur) > 0 Then
            out.Add CLng(cur)
            cur = ""
        End If
    Next i
    Set Numbers = out
End Function

Public Function FirstNumber(ByVal s As String) As Long
    Dim c As Collection: Set c = Numbers(s)
    If c.Count > 0 Then FirstNumber = c(1) Else FirstNumber = -1
End Function

' zero-pad every number (natural sort keys: AT-2 before AT-10)
Public Function NaturalPad(ByVal s As String) As String
    Dim i As Long, cur As String, out As String, ch As String
    s = UCase$(s)
    For i = 1 To Len(s) + 1
        If i <= Len(s) Then ch = Mid$(s, i, 1) Else ch = ""
        If IsDigitChar(ch) Then
            cur = cur & ch
        Else
            If Len(cur) > 0 Then out = out & Right$("0000" & cur, 4): cur = ""
            out = out & ch
        End If
    Next i
    NaturalPad = out
End Function

' ---- collections as sets/maps (no Dictionary on Mac) -----------------------
Public Function HasKey(ByVal c As Collection, ByVal key As String) As Boolean
    On Error GoTo No
    Dim v As Variant
    v = c(key)
    HasKey = True
    Exit Function
No:
    HasKey = False
End Function

Public Sub AddUnique(ByVal c As Collection, ByVal value As String)
    If Not HasKey(c, "k" & value) Then c.Add value, "k" & value
End Sub

' sort a 1-based Variant array of strings in place, by parallel sort keys
Public Sub SortByKeys(ByRef items As Variant, ByRef keys As Variant)
    Dim i As Long, j As Long, tk As Variant, ti As Variant
    For i = LBound(keys) + 1 To UBound(keys)
        tk = keys(i): ti = items(i)
        j = i - 1
        Do While j >= LBound(keys)
            If keys(j) <= tk Then Exit Do
            keys(j + 1) = keys(j): items(j + 1) = items(j)
            j = j - 1
        Loop
        keys(j + 1) = tk: items(j + 1) = ti
    Next i
End Sub

Public Function CollToArray(ByVal c As Collection) As Variant
    Dim a() As Variant, i As Long
    If c.Count = 0 Then CollToArray = Array(): Exit Function
    ReDim a(1 To c.Count)
    For i = 1 To c.Count: a(i) = c(i): Next i
    CollToArray = a
End Function

' ---- CSV (UTF-8) ---------------------------------------------------------
Public Function ReadUtf8File(ByVal path As String) As String
    Dim f As Integer, n As Long, b() As Byte
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then Close #f: Exit Function
    ReDim b(0 To n - 1)
    Get #f, , b
    Close #f
    ReadUtf8File = Utf8Decode(b)
End Function

Public Function Utf8Decode(b() As Byte) As String
    Dim i As Long, n As Long, cp As Long, out() As String, k As Long, chunk As String
    n = UBound(b) + 1
    i = 0
    If n >= 3 Then If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then i = 3   ' BOM
    ReDim out(0 To n \ 1024 + 1)
    Do While i < n
        If b(i) < &H80 Then
            cp = b(i): i = i + 1
        ElseIf (b(i) And &HE0) = &HC0 And i + 1 < n Then
            cp = (CLng(b(i) And &H1F) * &H40) Or (b(i + 1) And &H3F): i = i + 2
        ElseIf (b(i) And &HF0) = &HE0 And i + 2 < n Then
            cp = (CLng(b(i) And &HF) * &H1000) Or (CLng(b(i + 1) And &H3F) * &H40) Or (b(i + 2) And &H3F): i = i + 3
        ElseIf (b(i) And &HF8) = &HF0 And i + 3 < n Then
            cp = (CLng(b(i) And &H7) * &H40000) Or (CLng(b(i + 1) And &H3F) * &H1000) Or _
                 (CLng(b(i + 2) And &H3F) * &H40) Or (b(i + 3) And &H3F): i = i + 4
        Else
            cp = &HFFFD: i = i + 1
        End If
        If cp >= &H10000 Then
            cp = cp - &H10000
            chunk = chunk & ChrW(&HD800 + (cp \ &H400)) & ChrW(&HDC00 + (cp And &H3FF))
        Else
            chunk = chunk & ChrW(cp)
        End If
        If Len(chunk) >= 1024 Then out(k) = chunk: k = k + 1: chunk = ""
        If k > UBound(out) Then ReDim Preserve out(0 To k + 16)
    Loop
    out(k) = chunk
    ReDim Preserve out(0 To k)
    Utf8Decode = Join(out, "")
End Function

' RFC-4180-ish CSV -> 2D Variant (1-based rows/cols); quotes, "" and newlines in quotes
Public Function ParseCsv(ByVal text As String) As Variant
    Dim rows As New Collection, fields As Collection, i As Long, n As Long, ch As String
    Dim field As String, inQ As Boolean, maxCols As Long, r As Long, c As Long
    Dim parts() As String, pc As Long
    n = Len(text)
    Set fields = New Collection
    ReDim parts(0 To 63): pc = 0
    i = 1
    Do While i <= n
        ch = Mid$(text, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(text, i + 1, 1) = """" Then
                    field = field & """": i = i + 1
                Else
                    inQ = False
                End If
            Else
                field = field & ch
            End If
        Else
            Select Case ch
                Case """": inQ = True
                Case ",": fields.Add field: field = ""
                Case vbCr
                    ' handled with LF
                Case vbLf
                    fields.Add field: field = ""
                    rows.Add fields: Set fields = New Collection
                Case Else: field = field & ch
            End Select
        End If
        i = i + 1
    Loop
    If Len(field) > 0 Or fields.Count > 0 Then fields.Add field: rows.Add fields
    For r = 1 To rows.Count
        If rows(r).Count > maxCols Then maxCols = rows(r).Count
    Next r
    Dim out() As Variant
    If rows.Count = 0 Then ParseCsv = Empty: Exit Function
    ReDim out(1 To rows.Count, 1 To maxCols)
    For r = 1 To rows.Count
        For c = 1 To rows(r).Count
            out(r, c) = rows(r)(c)
        Next c
    Next r
    ParseCsv = out
End Function

' ---- parsing rules (same as the Python prototype / dashboard) --------------
Public Function IsHurdleYes(ByVal v As String) As Boolean
    v = LCase$(Trim$(v))
    If Len(v) = 0 Then Exit Function
    If InStr(v, "non") > 0 Then Exit Function
    IsHurdleYes = (InStr(v, "yes") > 0)
End Function

' headline weighting as a fraction (0.3), or -1 when unreadable
Public Function ShortWeight(ByVal text As String) As Double
    Dim t As String, p As Long, v As Double, e As Long, a As Double, b As Double, e2 As Long, prevDot As Boolean
    t = LCase$(Replace(text, "per cent", "percent"))
    ShortWeight = -1
    ' 1. "total = 50 percent"
    p = InStr(t, "total")
    Do While p > 0
        e = p + 5
        Do While e <= Len(t) And InStr(" :=", Mid$(t, e, 1)) > 0: e = e + 1: Loop
        If NumberAt(t, e, v, e2) Then
            If PercentAt(t, e2) Then ShortWeight = v / 100: Exit Function
        End If
        p = InStr(p + 1, t, "total")
    Loop
    ' 2. leading "4 x 5 percent"
    e = 1
    Do While e <= Len(t) And Mid$(t, e, 1) = " ": e = e + 1: Loop
    If NumberAt(t, e, a, e2) Then
        e = e2
        Do While e <= Len(t) And Mid$(t, e, 1) = " ": e = e + 1: Loop
        If e <= Len(t) And (Mid$(t, e, 1) = "x" Or Mid$(t, e, 1) = GTimes()) Then
            e = e + 1
            Do While e <= Len(t) And Mid$(t, e, 1) = " ": e = e + 1: Loop
            If NumberAt(t, e, b, e2) Then
                If PercentAt(t, e2) Then ShortWeight = a * b / 100: Exit Function
            End If
        End If
    End If
    ' 3. first "N percent" / "N%"
    For p = 1 To Len(t)
        If IsDigitChar(Mid$(t, p, 1)) And (p = 1 Or Not IsDigitChar(Mid$(t, IIf(p > 1, p - 1, 1), 1)) Or p = 1) Then
            If p = 1 Then
                prevDot = False
            Else
                prevDot = (Mid$(t, p - 1, 1) = ".")
            End If
            If Not prevDot Then
                If NumberAt(t, p, v, e2) Then
                    If PercentAt(t, e2) Then ShortWeight = v / 100: Exit Function
                End If
            End If
        End If
    Next p
End Function

Private Function NumberAt(ByVal t As String, ByVal p As Long, ByRef v As Double, ByRef endPos As Long) As Boolean
    Dim s As String, e As Long
    e = p
    Do While e <= Len(t)
        If IsDigitChar(Mid$(t, e, 1)) Then
            s = s & Mid$(t, e, 1)
        ElseIf Mid$(t, e, 1) = "." And Len(s) > 0 And InStr(s, ".") = 0 And e < Len(t) Then
            If IsDigitChar(Mid$(t, e + 1, 1)) Then s = s & "." Else Exit Do
        Else
            Exit Do
        End If
        e = e + 1
    Loop
    If Len(s) = 0 Then Exit Function
    v = Val(s)
    endPos = e
    NumberAt = True
End Function

Private Function PercentAt(ByVal t As String, ByVal p As Long) As Boolean
    Do While p <= Len(t) And Mid$(t, p, 1) = " ": p = p + 1: Loop
    If p > Len(t) Then Exit Function
    PercentAt = (Mid$(t, p, 1) = "%" Or Mid$(t, p, 6) = "percen")
End Function

Public Function ModeLabel(ByVal text As String) As String
    Dim t As String, ip As Boolean, ol As Boolean
    t = LCase$(text)
    ip = (InStr(t, "campus") Or InStr(t, "in person") Or InStr(t, "in-person") Or InStr(t, "field") Or _
          InStr(t, "placement") Or InStr(t, "site") Or InStr(t, "face to face") Or InStr(t, "face-to-face") Or InStr(t, "lab")) > 0
    ol = InStr(t, "online") > 0
    If ip And ol Then
        ModeLabel = "Both"
    ElseIf ip Then
        ModeLabel = "In person"
    ElseIf ol Then
        ModeLabel = "Online"
    Else
        ModeLabel = "Other"
    End If
End Function

Public Function ParticipationLabel(ByVal text As String) As String
    Dim t As String, ind As Boolean, grp As Boolean
    t = LCase$(text)
    ind = InStr(t, "individual") > 0
    grp = (InStr(t, "group") > 0 Or InStr(t, "team") > 0)
    If ind And grp Then
        ParticipationLabel = "Both"
    ElseIf ind Then
        ParticipationLabel = "Ind"
    ElseIf grp Then
        ParticipationLabel = "Grp"
    Else
        ParticipationLabel = GMdash()
    End If
End Function

Public Function AolTypeShort(ByVal text As String) As String
    Dim t As String
    t = Squash(text)
    Select Case LCase$(t)
        Case "", "na", "n/a", "not applicable", "no", "tbc", "not aol task": AolTypeShort = ""
        Case Else: AolTypeShort = Clip(t, 34)
    End Select
End Function

' alignment mark from a task row
Public Function MarkOf(ByVal alignmentMark As String, ByVal aol As String, ByVal hurdle As String) As String
    If Len(alignmentMark) > 0 Then MarkOf = alignmentMark: Exit Function
    Dim isAol As Boolean, isH As Boolean
    isAol = (LCase$(Trim$(aol)) = "yes")
    isH = IsHurdleYes(hurdle)
    If isAol And isH Then
        MarkOf = GFlag()
    ElseIf isAol Then
        MarkOf = GDot()
    ElseIf isH Then
        MarkOf = GDiamond()
    Else
        MarkOf = GRing()
    End If
End Function

' GLO cell vocabulary -> code + rank (Summary sheet)
Public Sub GloCode(ByVal v As String, ByRef code As String, ByRef rank As Long)
    Select Case LCase$(Trim$(v))
        Case "yes", "hurdle yes": code = "H": rank = 5
        Case "non-hurdle yes": code = "Y": rank = 4
        Case "data error": code = "E": rank = 3
        Case "tbc", "form not received": code = "?": rank = 2
        Case "no", "not assessed": code = GMid(): rank = 1
        Case Else: code = "": rank = 0
    End Select
End Sub

Public Function YearSortKey(ByVal y As String) As String
    Dim n As Long
    Select Case y
        Case "Non-major": n = 0
        Case "Y1": n = 1
        Case "Y2": n = 2
        Case "Y3": n = 3
        Case "Hons": n = 4
        Case "PG": n = 7
        Case "MRes": n = 8
        Case "PhD": n = 9
        Case Else: n = 99
    End Select
    YearSortKey = Format$(n, "00") & "|" & y
End Function

Public Function TriSortKey(ByVal t As String) As String
    Dim n As Long: n = FirstNumber(t)
    If n < 0 Then n = 9
    TriSortKey = Format$(n, "0") & "|" & t
End Function

Public Function YearLevelLabel(ByVal y As String) As String
    Select Case y
        Case "Hons": YearLevelLabel = "Y4 (Hons)"
        Case "PG": YearLevelLabel = "Y5 (PG)"
        Case Else: YearLevelLabel = y
    End Select
End Function

' Switch Excel events back on (click-to-filter needs them). Returns the state.
Public Function EventsOn() As String
    Application.EnableEvents = True
    EventsOn = "events " & Application.EnableEvents
End Function
