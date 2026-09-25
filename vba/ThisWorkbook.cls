Option Explicit

' ThisWorkbook document module
' Runs a light-touch initialisation when the workbook opens.

Private Sub Workbook_Open()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    If Not ws Is Nothing Then
        ws.Activate
        If Len(Trim(CStr(ws.Range("out_Status").Value))) = 0 Then
            ws.Range("out_Status").Value = "Ready. Enter inputs, then click Run Simulation. Use Refresh Data to import NSE files."
        End If
    End If
    On Error GoTo 0
End Sub
