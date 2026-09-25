Option Explicit

' modWebRefresh
' Optional DIRECT web download of Total Returns Index history from
' niftyindices.com (fallback when export files are not at hand).
'
' Requires: Windows Excel with internet access (uses MSXML2.XMLHTTP via WinINET
' so the session cookie set by the initial GET is reused for the POST).
' This is best-effort - corporate proxies / NSE bot-protection may block it.
' The file-based Refresh Data path remains the primary, most reliable method.

Private Const BASE As String = "https://www.niftyindices.com"
Private Const TRI_URL As String = "https://www.niftyindices.com/BackPage/getTotalReturnIndexString"

Public Sub WebRefresh()
    On Error GoTo EH
    Dim dash As Worksheet
    Set dash = ThisWorkbook.Worksheets("Dashboard")

    If UCase(Trim(CStr(ThisWorkbook.Worksheets("Config").Range("cfg_WebEnable").Value))) <> "YES" Then
        dash.Range("out_Status").Value = "WEB REFRESH DISABLED - set 'Enable web refresh' to Yes on the Config sheet first."
        Exit Sub
    End If

    dash.Range("out_Status").Value = "WEB REFRESH - contacting niftyindices.com ..."
    DoEvents

    Dim sd As String, ed As String
    sd = Format(CDate(ThisWorkbook.Worksheets("Config").Range("cfg_WebStart").Value), "dd-mmm-yyyy")
    ed = Format(CDate(ThisWorkbook.Worksheets("Config").Range("cfg_WebEnd").Value), "dd-mmm-yyyy")

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", BASE & "/reports/historical-data", False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.send

    Dim mapMom As Collection, mapVal As Collection, mapBen As Collection
    Dim serMom As Collection, serVal As Collection, serBen As Collection
    Set serMom = New Collection: Set serVal = New Collection: Set serBen = New Collection
    Set mapMom = FetchIndex(http, "NIFTY200 MOMENTUM 30", "Nifty200Momentm30", sd, ed, serMom)
    Set mapVal = FetchIndex(http, "Nifty500 Value 50", "Nifty500 Value 50", sd, ed, serVal)
    Set mapBen = FetchIndex(http, "NIFTY 500", "Nifty 500", sd, ed, serBen)

    If serMom.Count = 0 And serVal.Count = 0 And serBen.Count = 0 Then
        dash.Range("out_Status").Value = "WEB REFRESH FAILED - no data returned (check internet / proxy, or use file import)."
        Exit Sub
    End If

    ' aligned dates = present in all three
    Dim tmp() As Double, n As Long, i As Long, sVal As Double
    ReDim tmp(1 To serMom.Count + 1)
    n = 0
    For i = 1 To serMom.Count
        sVal = serMom(i)
        If Has(mapVal, sVal) And Has(mapBen, sVal) Then
            n = n + 1: tmp(n) = sVal
        End If
    Next i
    If n = 0 Then
        dash.Range("out_Status").Value = "WEB REFRESH FAILED - indices did not share common dates."
        Exit Sub
    End If
    ReDim Preserve tmp(1 To n)
    SortAsc tmp, 1, n

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Database_Daily")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Range("A2:D" & lastRow).ClearContents

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 4)
    For i = 1 To n
        outArr(i, 1) = CDate(tmp(i))
        outArr(i, 2) = ItemOf(mapMom, tmp(i))
        outArr(i, 3) = ItemOf(mapVal, tmp(i))
        outArr(i, 4) = ItemOf(mapBen, tmp(i))
    Next i
    ws.Range("A2").Resize(n, 4).Value = outArr
    ws.Columns("A:A").NumberFormat = "dd-mmm-yyyy"

    dash.Range("out_Status").Value = "WEB REFRESH COMPLETE - " & n & " daily rows downloaded from NSE (" & sd & " to " & ed & "). Now click Run Simulation."
    Exit Sub
EH:
    dash.Range("out_Status").Value = "WEB REFRESH ERROR - " & Err.Description & " (try the file-based Refresh Data instead)."
End Sub

' Returns a Collection of values keyed by CStr(serial); also fills 'serials'
' with the ordered list of date serials parsed.
Private Function FetchIndex(ByRef http As Object, ByVal nm As String, ByVal idx As String, _
                            ByVal sd As String, ByVal ed As String, ByRef serials As Collection) As Collection
    Dim c As New Collection
    Set FetchIndex = c

    Dim body As String
    body = "{""cinfo"":""{'name':'" & nm & "','startDate':'" & sd & "','endDate':'" & ed & "','indexName':'" & idx & "'}""}"

    http.Open "POST", TRI_URL, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.setRequestHeader "Referer", BASE & "/reports/historical-data"
    http.setRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.setRequestHeader "Origin", BASE
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.send body

    ParseTRI http.responseText, c, serials
End Function

' Minimal JSON scan: pull each ("Date","TotalReturnsIndex") pair.
Private Sub ParseTRI(ByVal json As String, ByRef c As Collection, ByRef serials As Collection)
    Dim p As Long, dStart As Long, dEnd As Long, vStart As Long, vEnd As Long
    Dim dtoken As String, vtoken As String
    Const DK As String = """Date"":"""
    Const VK As String = """TotalReturnsIndex"":"""

    p = InStr(1, json, DK)
    Do While p > 0
        dStart = p + Len(DK)
        dEnd = InStr(dStart, json, """")
        dtoken = Mid(json, dStart, dEnd - dStart)

        vStart = InStr(dEnd, json, VK)
        If vStart = 0 Then Exit Do
        vStart = vStart + Len(VK)
        vEnd = InStr(vStart, json, """")
        vtoken = Replace(Mid(json, vStart, vEnd - vStart), ",", "")

        Dim serial As Double
        If ParseTRIDate(dtoken, serial) And IsNumeric(vtoken) Then
            On Error Resume Next
            c.Add CDbl(vtoken), CStr(serial)
            If Err.Number = 0 Then serials.Add serial
            On Error GoTo 0
        End If

        p = InStr(vEnd, json, DK)
    Loop
End Sub

Private Function ParseTRIDate(ByVal s As String, ByRef outSerial As Double) As Boolean
    Dim parts() As String
    parts = Split(Application.WorksheetFunction.Trim(s), " ")
    If UBound(parts) <> 2 Then ParseTRIDate = False: Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(2)) Then ParseTRIDate = False: Exit Function
    Dim dd As Long, yy As Long, mm As Long
    dd = CLng(parts(0)): yy = CLng(parts(2))
    Select Case LCase(Left(parts(1), 3))
        Case "jan": mm = 1
        Case "feb": mm = 2
        Case "mar": mm = 3
        Case "apr": mm = 4
        Case "may": mm = 5
        Case "jun": mm = 6
        Case "jul": mm = 7
        Case "aug": mm = 8
        Case "sep": mm = 9
        Case "oct": mm = 10
        Case "nov": mm = 11
        Case "dec": mm = 12
        Case Else: ParseTRIDate = False: Exit Function
    End Select
    On Error Resume Next
    outSerial = CDbl(DateSerial(yy, mm, dd))
    ParseTRIDate = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function Has(ByRef c As Collection, ByVal serial As Double) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = c.Item(CStr(serial))
    Has = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ItemOf(ByRef c As Collection, ByVal serial As Double) As Double
    On Error Resume Next
    ItemOf = CDbl(c.Item(CStr(serial)))
    On Error GoTo 0
End Function

Private Sub SortAsc(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, tmp As Double
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < p: i = i + 1: Loop
        Do While a(j) > p: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortAsc a, lo, j
    If i < hi Then SortAsc a, i, hi
End Sub
