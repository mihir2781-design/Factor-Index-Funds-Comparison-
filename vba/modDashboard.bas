Option Explicit

' modDashboard
' Orchestrates the run: validate -> load -> compute -> publish. Also handles
' Reset. These three public Subs are wired to the Dashboard buttons.

Private Const R_AVG As Long = 18
Private Const R_WORST As Long = 19
Private Const R_BEST As Long = 20
Private Const R_P12 As Long = 21
Private Const R_P15 As Long = 22
Private Const R_CORPUS As Long = 23
' columns: Momentum=4(D) Value=5(E) Benchmark=6(F) Blended=7(G)

Public Sub RunSimulation()
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")

    Application.ScreenUpdating = False

    Dim msg As String
    If Not modValidation.ValidateInputs(msg) Then
        ws.Range("out_Status").Value = msg
        ClearResults
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If Not modRollingReturns.LoadData() Then
        ws.Range("out_Status").Value = "INSUFFICIENT DATA - the daily database is empty. Use Refresh Data or import NSE files first."
        ClearResults
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Dim mode As String, amount As Double, horizon As Long
    Dim sSerial As Double, eSerial As Double, wMom As Double, wVal As Double
    mode = Trim(CStr(ws.Range("in_Mode").Value))
    amount = CDbl(ws.Range("in_Amount").Value)
    horizon = CLng(ws.Range("in_Horizon").Value)
    sSerial = CDbl(CDate(ws.Range("in_StartDate").Value))
    eSerial = CDbl(CDate(ws.Range("in_EndDate").Value))
    wMom = CDbl(ws.Range("in_AllocMom").Value) / 100#
    wVal = CDbl(ws.Range("in_AllocVal").Value) / 100#

    Dim mMom As TMetrics, mVal As TMetrics
    Dim mBen As TMetrics, mBlend As TMetrics
    Dim nWin As Long

    If Not modRollingReturns.RunEngine(mode, amount, horizon, sSerial, eSerial, wMom, wVal, _
                                       mMom, mVal, mBen, mBlend, nWin) Then
        ws.Range("out_Status").Value = "INSUFFICIENT DATA - no complete " & horizon & "-year rolling window fits inside the selected sample range."
        ClearResults
        Application.ScreenUpdating = True
        Exit Sub
    End If

    WriteColumn ws, 4, mMom
    WriteColumn ws, 5, mVal
    WriteColumn ws, 6, mBen
    WriteColumn ws, 7, mBlend

    ws.Range("B25").Value = nWin
    ws.Range("B26").Value = Now

    ws.Range("out_Status").Value = "SIMULATION COMPLETE - " & nWin & " rolling windows evaluated (" _
        & UCase(mode) & ", " & horizon & "Y horizon). Blended = " _
        & Format(wMom, "0%") & " Momentum / " & Format(wVal, "0%") & " Value."

    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    ws.Range("out_Status").Value = "RUN ERROR - " & Err.Description
End Sub

Private Sub WriteColumn(ByRef ws As Worksheet, ByVal col As Long, ByRef m As TMetrics)
    ws.Cells(R_AVG, col).Value = m.Avg
    ws.Cells(R_WORST, col).Value = m.Worst
    ws.Cells(R_BEST, col).Value = m.Best
    ws.Cells(R_P12, col).Value = m.P12
    ws.Cells(R_P15, col).Value = m.P15
    ws.Cells(R_CORPUS, col).Value = m.Corpus
End Sub

Private Sub ClearResults()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range(ws.Cells(R_AVG, 4), ws.Cells(R_CORPUS, 7)).ClearContents
    ws.Range("B25").ClearContents
    ws.Range("B26").ClearContents
End Sub

Public Sub ResetInputs()
    On Error GoTo EH
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    ws.Range("in_Mode").Value = "SIP"
    ws.Range("in_Amount").Value = 25000
    ws.Range("in_Horizon").Value = 5
    ws.Range("in_StartDate").Value = DateSerial(2012, 1, 1)
    ws.Range("in_EndDate").Value = DateSerial(2024, 12, 31)
    ws.Range("in_AllocMom").Value = 60
    ws.Range("in_AllocVal").Value = 40
    ClearResults
    ws.Range("out_Status").Value = "Inputs reset to defaults. Click Run Simulation to compute."
    Exit Sub
EH:
    ws.Range("out_Status").Value = "RESET ERROR - " & Err.Description
End Sub
