Option Explicit

' modRollingReturns
' Loads the daily database into memory, generates valid rolling windows
' (monthly cadence) within the user-selected sample period, and computes
' asset-wise + blended metrics.
'   Lumpsum -> rolling CAGR
'   SIP      -> exact rolling XIRR (Newton-Raphson, VBA-native)

' ---- In-memory database (1-based arrays) ----
Public gN As Long
Public gDates() As Double        ' Excel date serials, ascending
Public gMom() As Double          ' Nifty 200 Momentum 30 TRI
Public gVal() As Double          ' Nifty 500 Value 50 TRI
Public gBen() As Double          ' Nifty 500 Parent Benchmark TRI
Public gMonthStart() As Long     ' indices that are the first trading row of their month
Public gMS As Long               ' count of month-start rows

Public Type TMetrics
    Avg As Double
    Worst As Double
    Best As Double
    P12 As Double
    P15 As Double
    Corpus As Double
    Count As Long
End Type

' Load Database_Daily into module arrays. Returns False if there is no data.
Public Function LoadData() As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Database_Daily")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        gN = 0: LoadData = False: Exit Function
    End If

    Dim raw As Variant
    raw = ws.Range("A2:D" & lastRow).Value

    Dim rows As Long
    rows = lastRow - 1

    ReDim gDates(1 To rows)
    ReDim gMom(1 To rows)
    ReDim gVal(1 To rows)
    ReDim gBen(1 To rows)

    Dim i As Long, k As Long
    k = 0
    For i = 1 To rows
        If IsDate(raw(i, 1)) And IsNumeric(raw(i, 2)) And IsNumeric(raw(i, 3)) And IsNumeric(raw(i, 4)) Then
            k = k + 1
            gDates(k) = CDbl(CDate(raw(i, 1)))
            gMom(k) = CDbl(raw(i, 2))
            gVal(k) = CDbl(raw(i, 3))
            gBen(k) = CDbl(raw(i, 4))
        End If
    Next i

    gN = k
    If gN = 0 Then LoadData = False: Exit Function

    ReDim Preserve gDates(1 To gN)
    ReDim Preserve gMom(1 To gN)
    ReDim Preserve gVal(1 To gN)
    ReDim Preserve gBen(1 To gN)

    BuildMonthStarts
    LoadData = True
End Function

Private Sub BuildMonthStarts()
    ReDim gMonthStart(1 To gN)
    gMS = 0
    Dim i As Long, prevKey As String, curKey As String
    prevKey = ""
    For i = 1 To gN
        curKey = Format(CDate(gDates(i)), "yyyy-mm")
        If curKey <> prevKey Then
            gMS = gMS + 1
            gMonthStart(gMS) = i
            prevKey = curKey
        End If
    Next i
    If gMS > 0 Then ReDim Preserve gMonthStart(1 To gMS)
End Sub

' Binary search: first index in [lo,hi] with gDates(idx) >= target, else -1.
Private Function LowerBound(ByVal target As Double, ByVal lo As Long, ByVal hi As Long) As Long
    Dim l As Long, h As Long, mid As Long, res As Long
    l = lo: h = hi: res = -1
    Do While l <= h
        mid = (l + h) \ 2
        If gDates(mid) >= target Then
            res = mid: h = mid - 1
        Else
            l = mid + 1
        End If
    Loop
    LowerBound = res
End Function

' ---- Core engine ----
' Fills metrics for Momentum(1), Value(2), Benchmark(3), Blended(4).
' Returns False when no valid rolling window exists (INSUFFICIENT DATA).
Public Function RunEngine(ByVal mode As String, ByVal amount As Double, ByVal horizon As Long, _
                          ByVal startSerial As Double, ByVal endSerial As Double, _
                          ByVal wMom As Double, ByVal wVal As Double, _
                          ByRef mMom As TMetrics, ByRef mVal As TMetrics, _
                          ByRef mBen As TMetrics, ByRef mBlend As TMetrics, _
                          ByRef nWindows As Long) As Boolean

    Dim isSIP As Boolean
    isSIP = (UCase(mode) = "SIP")

    ' sample bounds
    Dim firstIdx As Long, lastIdx As Long
    firstIdx = LowerBound(startSerial, 1, gN)
    If firstIdx = -1 Then RunEngine = False: Exit Function
    ' last index with date <= endSerial
    lastIdx = LowerBound(endSerial + 0.5, 1, gN)
    If lastIdx = -1 Then
        lastIdx = gN
    Else
        lastIdx = lastIdx - 1
    End If
    If lastIdx < firstIdx Then RunEngine = False: Exit Function

    Dim maxW As Long
    maxW = gMS + 1
    Dim rMom() As Double, rVal() As Double, rBen() As Double, rBlend() As Double
    ReDim rMom(1 To maxW): ReDim rVal(1 To maxW)
    ReDim rBen(1 To maxW): ReDim rBlend(1 To maxW)

    Dim k As Long
    k = 0

    Dim j As Long, e As Long, mIdx As Long
    Dim entrySerial As Double, matTarget As Double, entryDate As Date

    For j = 1 To gMS
        e = gMonthStart(j)
        If e >= firstIdx And e <= lastIdx Then
            entrySerial = gDates(e)
            entryDate = CDate(entrySerial)
            matTarget = CDbl(DateAdd("yyyy", horizon, entryDate))
            mIdx = LowerBound(matTarget, e, gN)
            If mIdx <> -1 Then
                If mIdx > e And gDates(mIdx) <= endSerial Then
                    k = k + 1
                    If isSIP Then
                        rMom(k) = SIPReturn(gMom, e, mIdx, matTarget, entrySerial, amount, 1#, 0#, gVal)
                        rVal(k) = SIPReturn(gVal, e, mIdx, matTarget, entrySerial, amount, 1#, 0#, gMom)
                        rBen(k) = SIPReturn(gBen, e, mIdx, matTarget, entrySerial, amount, 1#, 0#, gBen)
                        rBlend(k) = SIPReturnBlend(e, mIdx, matTarget, entrySerial, amount, wMom, wVal)
                    Else
                        rMom(k) = (gMom(mIdx) / gMom(e)) ^ (1# / horizon) - 1#
                        rVal(k) = (gVal(mIdx) / gVal(e)) ^ (1# / horizon) - 1#
                        rBen(k) = (gBen(mIdx) / gBen(e)) ^ (1# / horizon) - 1#
                        Dim gf As Double
                        gf = modBlendEngine.BlendedGrowthFactor(gMom(e), gMom(mIdx), gVal(e), gVal(mIdx), wMom, wVal)
                        rBlend(k) = gf ^ (1# / horizon) - 1#
                    End If
                    CacheWindow k, entryDate, CDate(gDates(mIdx)), rMom(k), rVal(k), rBen(k), rBlend(k)
                End If
            End If
        End If
    Next j

    nWindows = k
    If k = 0 Then RunEngine = False: Exit Function

    mMom = Summarise(rMom, k, amount, horizon, isSIP)
    mVal = Summarise(rVal, k, amount, horizon, isSIP)
    mBen = Summarise(rBen, k, amount, horizon, isSIP)
    mBlend = Summarise(rBlend, k, amount, horizon, isSIP)

    RunEngine = True
End Function

' SIP XIRR for a single asset price series.
' wA/wB let us reuse this for pure single-asset (wA=1, wB=0); the second series
' is ignored when wB = 0.
Private Function SIPReturn(ByRef price() As Double, ByVal e As Long, ByVal mIdx As Long, _
                           ByVal matTarget As Double, ByVal entrySerial As Double, _
                           ByVal amount As Double, ByVal wA As Double, ByVal wB As Double, _
                           ByRef priceB() As Double) As Double
    Dim cf() As Double, t() As Double
    Dim n As Long
    n = 0
    ReDim cf(1 To gMS + 1): ReDim t(1 To gMS + 1)

    Dim units As Double
    units = 0#
    Dim j As Long, c As Long
    For j = 1 To gMS
        c = gMonthStart(j)
        If gDates(c) >= entrySerial And gDates(c) < gDates(mIdx) Then
            n = n + 1
            cf(n) = -amount
            t(n) = (gDates(c) - entrySerial) / 365#
            units = units + amount / price(c)
        End If
    Next j

    n = n + 1
    cf(n) = units * price(mIdx)
    t(n) = (gDates(mIdx) - entrySerial) / 365#

    SIPReturn = XIRR(cf, t, n)
End Function

' SIP XIRR for the blended portfolio (fixed split at every contribution).
Private Function SIPReturnBlend(ByVal e As Long, ByVal mIdx As Long, ByVal matTarget As Double, _
                                ByVal entrySerial As Double, ByVal amount As Double, _
                                ByVal wMom As Double, ByVal wVal As Double) As Double
    Dim cf() As Double, t() As Double
    Dim n As Long
    n = 0
    ReDim cf(1 To gMS + 1): ReDim t(1 To gMS + 1)

    Dim unitsMom As Double, unitsVal As Double, amtMom As Double, amtVal As Double
    unitsMom = 0#: unitsVal = 0#
    Dim j As Long, c As Long
    For j = 1 To gMS
        c = gMonthStart(j)
        If gDates(c) >= entrySerial And gDates(c) < gDates(mIdx) Then
            n = n + 1
            cf(n) = -amount
            t(n) = (gDates(c) - entrySerial) / 365#
            modBlendEngine.ContributionSplit amount, wMom, wVal, amtMom, amtVal
            unitsMom = unitsMom + amtMom / gMom(c)
            unitsVal = unitsVal + amtVal / gVal(c)
        End If
    Next j

    n = n + 1
    cf(n) = modBlendEngine.BlendedSIPMaturity(unitsMom, unitsVal, gMom(mIdx), gVal(mIdx))
    t(n) = (gDates(mIdx) - entrySerial) / 365#

    SIPReturnBlend = XIRR(cf, t, n)
End Function

Private Function Summarise(ByRef r() As Double, ByVal n As Long, ByVal amount As Double, _
                           ByVal horizon As Long, ByVal isSIP As Boolean) As TMetrics
    Dim res As TMetrics
    Dim i As Long, s As Double, c12 As Long, c15 As Long
    res.Worst = r(1): res.Best = r(1): s = 0#
    For i = 1 To n
        s = s + r(i)
        If r(i) < res.Worst Then res.Worst = r(i)
        If r(i) > res.Best Then res.Best = r(i)
        If r(i) > 0.12 Then c12 = c12 + 1
        If r(i) > 0.15 Then c15 = c15 + 1
    Next i
    res.Avg = s / n
    res.P12 = c12 / n
    res.P15 = c15 / n
    res.Count = n
    res.Corpus = ProjectCorpus(res.Avg, amount, horizon, isSIP)
    Summarise = res
End Function

' Final projected maturity corpus from the average rolling return.
Private Function ProjectCorpus(ByVal avgRet As Double, ByVal amount As Double, _
                               ByVal horizon As Long, ByVal isSIP As Boolean) As Double
    If isSIP Then
        Dim mth As Double, nMonths As Long
        mth = (1 + avgRet) ^ (1# / 12#) - 1#
        nMonths = horizon * 12
        If Abs(mth) < 0.0000000001 Then
            ProjectCorpus = amount * nMonths
        Else
            ' contributions at the start of each month
            ProjectCorpus = amount * (((1 + mth) ^ nMonths - 1) / mth) * (1 + mth)
        End If
    Else
        ProjectCorpus = amount * (1 + avgRet) ^ horizon
    End If
End Function

' ---- XIRR (Newton-Raphson with bisection fallback) ----
Public Function XIRR(ByRef cf() As Double, ByRef t() As Double, ByVal n As Long) As Double
    Dim r As Double, i As Long, f As Double, df As Double, base As Double, rn As Double, it As Long
    r = 0.1
    For it = 1 To 100
        f = 0#: df = 0#
        For i = 1 To n
            base = 1 + r
            If base <= 0.000001 Then base = 0.000001
            f = f + cf(i) / (base ^ t(i))
            df = df - cf(i) * t(i) / (base ^ (t(i) + 1))
        Next i
        If Abs(df) < 0.0000000001 Then Exit For
        rn = r - f / df
        If rn < -0.9999 Then rn = -0.9999
        If Abs(rn - r) < 0.0000001 Then r = rn: Exit For
        r = rn
    Next it

    If r <= -0.9999 Or r > 100 Then r = XIRR_Bisect(cf, t, n)
    XIRR = r
End Function

Private Function NPVr(ByRef cf() As Double, ByRef t() As Double, ByVal n As Long, ByVal r As Double) As Double
    Dim i As Long, s As Double, base As Double
    base = 1 + r
    If base <= 0.000001 Then base = 0.000001
    For i = 1 To n
        s = s + cf(i) / (base ^ t(i))
    Next i
    NPVr = s
End Function

Private Function XIRR_Bisect(ByRef cf() As Double, ByRef t() As Double, ByVal n As Long) As Double
    Dim lo As Double, hi As Double, mid As Double, fLo As Double, fMid As Double, it As Long
    lo = -0.99: hi = 5#
    fLo = NPVr(cf, t, n, lo)
    If fLo * NPVr(cf, t, n, hi) > 0 Then
        XIRR_Bisect = 0#: Exit Function
    End If
    For it = 1 To 200
        mid = (lo + hi) / 2#
        fMid = NPVr(cf, t, n, mid)
        If Abs(fMid) < 0.0000001 Then Exit For
        If fLo * fMid < 0 Then
            hi = mid
        Else
            lo = mid: fLo = fMid
        End If
    Next it
    XIRR_Bisect = mid
End Function

' Store per-window detail on Calc_Cache for auditing/debugging.
Private Sub CacheWindow(ByVal k As Long, ByVal entryDate As Date, ByVal matDate As Date, _
                        ByVal rMom As Double, ByVal rVal As Double, ByVal rBen As Double, ByVal rBlend As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Calc_Cache")
    If k = 1 Then
        ws.Range("A2:F100000").ClearContents
    End If
    Dim rw As Long
    rw = k + 1
    ws.Cells(rw, 1).Value = entryDate
    ws.Cells(rw, 2).Value = matDate
    ws.Cells(rw, 3).Value = rMom
    ws.Cells(rw, 4).Value = rVal
    ws.Cells(rw, 5).Value = rBen
    ws.Cells(rw, 6).Value = rBlend
End Sub
