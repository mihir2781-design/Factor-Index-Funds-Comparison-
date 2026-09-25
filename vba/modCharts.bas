Option Explicit

' modCharts
' Populates the hidden Charts_Data sheet after a run so the Dashboard's
' Equity Curve and Rolling-Return Distribution charts refresh automatically.

Public Sub BuildChartData(ByVal startSerial As Double, ByVal endSerial As Double, _
                          ByVal wMom As Double, ByVal wVal As Double, ByVal nWindows As Long)
    Dim cd As Worksheet
    Set cd = ThisWorkbook.Worksheets("Charts_Data")
    cd.Range("A2:E100000").ClearContents
    cd.Range("I2:L13").ClearContents

    ' ---- Equity curve (monthly, rebased to 100 at sample start) ----
    Dim baseIdx As Long, j As Long, e As Long, rw As Long
    Dim mBase As Double, vBase As Double, bBase As Double
    baseIdx = -1
    For j = 1 To modRollingReturns.gMS
        e = modRollingReturns.gMonthStart(j)
        If modRollingReturns.gDates(e) >= startSerial And modRollingReturns.gDates(e) <= endSerial Then
            baseIdx = e: Exit For
        End If
    Next j

    If baseIdx <> -1 Then
        mBase = modRollingReturns.gMom(baseIdx)
        vBase = modRollingReturns.gVal(baseIdx)
        bBase = modRollingReturns.gBen(baseIdx)
        rw = 1
        For j = 1 To modRollingReturns.gMS
            e = modRollingReturns.gMonthStart(j)
            If modRollingReturns.gDates(e) >= startSerial And modRollingReturns.gDates(e) <= endSerial Then
                rw = rw + 1
                cd.Cells(rw, 1).Value = CDate(modRollingReturns.gDates(e))
                cd.Cells(rw, 2).Value = 100# * modRollingReturns.gMom(e) / mBase
                cd.Cells(rw, 3).Value = 100# * modRollingReturns.gVal(e) / vBase
                cd.Cells(rw, 4).Value = 100# * modRollingReturns.gBen(e) / bBase
                cd.Cells(rw, 5).Value = wMom * 100# * (modRollingReturns.gMom(e) / mBase) _
                                      + wVal * 100# * (modRollingReturns.gVal(e) / vBase)
            End If
        Next j
    End If

    ' ---- Rolling-return distribution (histogram from Calc_Cache) ----
    Dim cc As Worksheet
    Set cc = ThisWorkbook.Worksheets("Calc_Cache")
    Dim countsMom(0 To 11) As Long, countsVal(0 To 11) As Long
    Dim countsBen(0 To 11) As Long, countsBlend(0 To 11) As Long

    Dim i As Long
    If nWindows > 0 Then
        Dim data As Variant
        data = cc.Range("C2:F" & (nWindows + 1)).Value
        For i = 1 To nWindows
            AddToBin countsMom, CDbl(data(i, 1))
            AddToBin countsVal, CDbl(data(i, 2))
            AddToBin countsBen, CDbl(data(i, 3))
            AddToBin countsBlend, CDbl(data(i, 4))
        Next i
    End If

    For i = 0 To 11
        cd.Cells(i + 2, 9).Value = countsMom(i)
        cd.Cells(i + 2, 10).Value = countsVal(i)
        cd.Cells(i + 2, 11).Value = countsBen(i)
        cd.Cells(i + 2, 12).Value = countsBlend(i)
    Next i
End Sub

Private Sub AddToBin(ByRef counts() As Long, ByVal r As Double)
    Dim idx As Long
    If r < -0.1 Then
        idx = 0
    ElseIf r >= 0.4 Then
        idx = 11
    Else
        idx = 3 + Int(r / 0.05)
    End If
    If idx < 0 Then idx = 0
    If idx > 11 Then idx = 11
    counts(idx) = counts(idx) + 1
End Sub
