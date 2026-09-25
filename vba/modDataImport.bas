Option Explicit

' modDataImport
' File-based import of official NSE / NSE Indices CSV exports.
' Reads the three configured export files, auto-detects the Date and TRI/value
' columns, merges everything into one chronological daily table, forward-fills
' gaps, and writes the result to Database_Daily.
'
' If no files are configured (or none can be found) the existing seeded data
' is retained and a clear status message is shown - nothing is destroyed.

Public Sub RefreshData()
    On Error GoTo EH
    Dim dash As Worksheet
    Set dash = ThisWorkbook.Worksheets("Dashboard")
    dash.Range("out_Status").Value = "Refreshing data from NSE export files..."

    Dim pMom As String, pVal As String, pBen As String
    pMom = Trim(CStr(ThisWorkbook.Worksheets("Config").Range("cfg_MomFile").Value))
    pVal = Trim(CStr(ThisWorkbook.Worksheets("Config").Range("cfg_ValFile").Value))
    pBen = Trim(CStr(ThisWorkbook.Worksheets("Config").Range("cfg_BenFile").Value))

    Dim haveAny As Boolean
    haveAny = (Len(pMom) > 0 And FileExists(pMom)) _
           Or (Len(pVal) > 0 And FileExists(pVal)) _
           Or (Len(pBen) > 0 And FileExists(pBen))

    If Not haveAny Then
        dash.Range("out_Status").Value = "REFRESH SKIPPED - no import files found. Set file paths on the Config sheet. Existing data retained."
        Exit Sub
    End If

    ' Per-series maps: key = date serial (as string), item = value
    Dim mapMom As Collection, mapVal As Collection, mapBen As Collection
    Dim allDates As Collection
    Set allDates = New Collection

    Set mapMom = LoadSeries(pMom, allDates)
    Set mapVal = LoadSeries(pVal, allDates)
    Set mapBen = LoadSeries(pBen, allDates)

    ' Sort the unique dates ascending
    Dim n As Long, i As Long, j As Long
    n = allDates.Count
    If n = 0 Then
        dash.Range("out_Status").Value = "REFRESH FAILED - no valid rows parsed from the export files."
        Exit Sub
    End If
    Dim ds() As Double
    ReDim ds(1 To n)
    For i = 1 To n
        ds(i) = allDates(i)
    Next i
    QuickSort ds, 1, n

    ' Build the merged table with forward-fill
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Database_Daily")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Range("A2:D" & lastRow).ClearContents

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 4)

    Dim lastMom As Double, lastVal As Double, lastBen As Double
    Dim haveMom As Boolean, haveVal As Boolean, haveBen As Boolean
    Dim rowsWritten As Long
    rowsWritten = 0

    For i = 1 To n
        Dim key As String, v As Double
        key = CStr(ds(i))

        If TryGet(mapMom, key, v) Then lastMom = v: haveMom = True
        If TryGet(mapVal, key, v) Then lastVal = v: haveVal = True
        If TryGet(mapBen, key, v) Then lastBen = v: haveBen = True

        ' only start writing once every configured series has a value
        If (Len(pMom) = 0 Or haveMom) And (Len(pVal) = 0 Or haveVal) And (Len(pBen) = 0 Or haveBen) Then
            rowsWritten = rowsWritten + 1
            outArr(rowsWritten, 1) = CDate(ds(i))
            outArr(rowsWritten, 2) = IIf(haveMom, lastMom, CVErr(2042))
            outArr(rowsWritten, 3) = IIf(haveVal, lastVal, CVErr(2042))
            outArr(rowsWritten, 4) = IIf(haveBen, lastBen, CVErr(2042))
        End If
    Next i

    If rowsWritten > 0 Then
        ws.Range("A2").Resize(rowsWritten, 4).Value = ResizeOut(outArr, rowsWritten)
        ws.Columns("A:A").NumberFormat = "dd-mmm-yyyy"
    End If

    dash.Range("out_Status").Value = "REFRESH COMPLETE - " & rowsWritten & " daily rows loaded and normalised. Now click Run Simulation."
    Exit Sub
EH:
    dash.Range("out_Status").Value = "REFRESH ERROR - " & Err.Description
End Sub

Private Function ResizeOut(ByRef src() As Variant, ByVal rows As Long) As Variant
    Dim o() As Variant, i As Long, c As Long
    ReDim o(1 To rows, 1 To 4)
    For i = 1 To rows
        For c = 1 To 4
            o(i, c) = src(i, c)
        Next c
    Next i
    ResizeOut = o
End Function

' Parse one CSV export into a date->value Collection and register dates.
Private Function LoadSeries(ByVal path As String, ByRef allDates As Collection) As Collection
    Dim map As Collection
    Set map = New Collection
    Set LoadSeries = map
    If Len(path) = 0 Or Not FileExists(path) Then Exit Function

    Dim ff As Integer, line As String, first As Boolean
    Dim dateCol As Long, valCol As Long
    first = True
    dateCol = -1: valCol = -1

    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, line
        If Len(Trim(line)) = 0 Then GoTo ContinueLoop
        Dim parts() As String
        parts = SplitCSV(line)

        If first Then
            DetectColumns parts, dateCol, valCol
            first = False
        Else
            If dateCol >= 0 And valCol >= 0 And UBound(parts) >= valCol Then
                Dim d As Double, ok As Boolean
                ok = ParseAnyDate(CleanCell(parts(dateCol)), d)
                If ok Then
                    Dim raw As String, val As Double
                    raw = CleanCell(parts(valCol))
                    If IsNumeric(raw) Then
                        val = CDbl(raw)
                        Dim key As String
                        key = CStr(d)
                        On Error Resume Next
                        map.Add val, key
                        allDates.Add d, key   ' key enforces uniqueness across series
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
ContinueLoop:
    Loop
    Close #ff
End Function

Private Sub DetectColumns(ByRef hdr() As String, ByRef dateCol As Long, ByRef valCol As Long)
    Dim i As Long, h As String
    dateCol = -1: valCol = -1
    ' Date column
    For i = 0 To UBound(hdr)
        h = LCase(CleanCell(hdr(i)))
        If InStr(h, "date") > 0 Then dateCol = i: Exit For
    Next i
    If dateCol = -1 Then dateCol = 0
    ' Value column by priority
    Dim pri As Variant, p As Variant
    pri = Array("total return", "tri", "total returns index", "close", "index value", "value")
    For Each p In pri
        For i = 0 To UBound(hdr)
            If i <> dateCol Then
                h = LCase(CleanCell(hdr(i)))
                If InStr(h, CStr(p)) > 0 Then valCol = i: Exit For
            End If
        Next i
        If valCol >= 0 Then Exit For
    Next p
    ' fallback: first non-date column
    If valCol = -1 And UBound(hdr) >= 1 Then
        valCol = IIf(dateCol = 0, 1, 0)
    End If
End Sub

' Robust date parse for common NSE export formats.
Private Function ParseAnyDate(ByVal s As String, ByRef outSerial As Double) As Boolean
    s = Trim(s)
    If Len(s) = 0 Then ParseAnyDate = False: Exit Function

    ' Try native first
    On Error Resume Next
    Dim d As Date
    d = CDate(s)
    If Err.Number = 0 Then
        outSerial = CDbl(d): ParseAnyDate = True: On Error GoTo 0: Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ' Normalise separators and try "dd MMM yyyy"
    Dim t As String
    t = Replace(Replace(s, "-", " "), "/", " ")
    Dim parts() As String
    parts = Split(Application.WorksheetFunction.Trim(t), " ")
    If UBound(parts) = 2 Then
        Dim dd As Long, yy As Long, mm As Long
        If IsNumeric(parts(0)) And IsNumeric(parts(2)) Then
            dd = CLng(parts(0)): yy = CLng(parts(2))
            mm = MonthFromName(parts(1))
            If mm = 0 And IsNumeric(parts(1)) Then mm = CLng(parts(1))
            If mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 And yy > 1900 Then
                On Error Resume Next
                outSerial = CDbl(DateSerial(yy, mm, dd))
                If Err.Number = 0 Then ParseAnyDate = True Else ParseAnyDate = False
                On Error GoTo 0
                Exit Function
            End If
        End If
    End If
    ParseAnyDate = False
End Function

Private Function MonthFromName(ByVal s As String) As Long
    Dim m As String
    m = LCase(Left(Trim(s), 3))
    Select Case m
        Case "jan": MonthFromName = 1
        Case "feb": MonthFromName = 2
        Case "mar": MonthFromName = 3
        Case "apr": MonthFromName = 4
        Case "may": MonthFromName = 5
        Case "jun": MonthFromName = 6
        Case "jul": MonthFromName = 7
        Case "aug": MonthFromName = 8
        Case "sep": MonthFromName = 9
        Case "oct": MonthFromName = 10
        Case "nov": MonthFromName = 11
        Case "dec": MonthFromName = 12
        Case Else: MonthFromName = 0
    End Select
End Function

Private Function CleanCell(ByVal s As String) As String
    s = Replace(s, Chr(34), "")
    s = Replace(s, ",", "")   ' strip thousands separators inside numbers/quoted cells
    CleanCell = Trim(s)
End Function

' CSV split that respects simple double-quoted fields.
Private Function SplitCSV(ByVal line As String) As String()
    Dim res() As String, n As Long, i As Long, ch As String
    Dim inQ As Boolean, cur As String
    ReDim res(0 To 0)
    n = 0: inQ = False: cur = ""
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If ch = Chr(34) Then
            inQ = Not inQ
        ElseIf ch = "," And Not inQ Then
            res(n) = cur: cur = ""
            n = n + 1: ReDim Preserve res(0 To n)
        Else
            cur = cur & ch
        End If
    Next i
    res(n) = cur
    SplitCSV = res
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(path)) > 0)
    On Error GoTo 0
End Function

Private Function TryGet(ByRef c As Collection, ByVal key As String, ByRef outVal As Double) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = c.Item(key)
    If Err.Number = 0 Then
        outVal = CDbl(v): TryGet = True
    Else
        TryGet = False
    End If
    On Error GoTo 0
End Function

Private Sub QuickSort(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, tmp As Double
    i = lo: j = hi
    p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSort a, lo, j
    If i < hi Then QuickSort a, i, hi
End Sub
