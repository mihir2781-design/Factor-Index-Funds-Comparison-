Option Explicit

' modValidation
' Hard input validation. Never silently corrects values.
' Returns True only when every rule passes; otherwise writes an explicit
' message to the Dashboard status cell and returns False.

Public Function ValidateInputs(ByRef outMsg As String) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")

    Dim modeRaw As String, horizonRaw As Variant, amt As Variant
    Dim sdt As Variant, edt As Variant, wm As Variant, wv As Variant

    modeRaw = Trim(CStr(ws.Range("in_Mode").Value))
    horizonRaw = ws.Range("in_Horizon").Value
    amt = ws.Range("in_Amount").Value
    sdt = ws.Range("in_StartDate").Value
    edt = ws.Range("in_EndDate").Value
    wm = ws.Range("in_AllocMom").Value
    wv = ws.Range("in_AllocVal").Value

    ' --- Mode ---
    Dim m As String
    m = UCase(modeRaw)
    If m <> "SIP" AND m <> "LUMPSUM" Then
        outMsg = "INVALID MODE - enter exactly 'SIP' or 'Lumpsum'."
        ValidateInputs = False: Exit Function
    End If

    ' --- Amount ---
    If Not IsNumeric(amt) Then
        outMsg = "INVALID AMOUNT - enter a positive number.": ValidateInputs = False: Exit Function
    End If
    If CDbl(amt) <= 0 Then
        outMsg = "INVALID AMOUNT - amount must be greater than 0.": ValidateInputs = False: Exit Function
    End If

    ' --- Horizon ---
    If Not IsNumeric(horizonRaw) Then
        outMsg = "INVALID HORIZON - enter 3, 5, or 10.": ValidateInputs = False: Exit Function
    End If
    Dim h As Long
    h = CLng(horizonRaw)
    If h <> 3 AND h <> 5 AND h <> 10 Then
        outMsg = "INVALID HORIZON - enter 3, 5, or 10.": ValidateInputs = False: Exit Function
    End If

    ' --- Allocation weights ---
    If Not IsNumeric(wm) Or Not IsNumeric(wv) Then
        outMsg = "INVALID ALLOCATION - Momentum % and Value % must be numbers.": ValidateInputs = False: Exit Function
    End If
    If CDbl(wm) < 0 Or CDbl(wv) < 0 Then
        outMsg = "INVALID ALLOCATION - weights cannot be negative.": ValidateInputs = False: Exit Function
    End If
    If CDbl(wm) + CDbl(wv) <> 100 Then
        outMsg = "INVALID ALLOCATION - Momentum % + Value % must equal exactly 100 (currently " _
                 & Format(CDbl(wm) + CDbl(wv), "0.##") & ").": ValidateInputs = False: Exit Function
    End If

    ' --- Dates ---
    If Not IsDate(sdt) Or Not IsDate(edt) Then
        outMsg = "INVALID DATES - Start Date and End Date must be valid dates.": ValidateInputs = False: Exit Function
    End If
    If CDate(sdt) >= CDate(edt) Then
        outMsg = "INVALID DATES - Start Date must be before End Date.": ValidateInputs = False: Exit Function
    End If

    ' --- Date span sufficiency vs horizon ---
    If DateAdd("yyyy", h, CDate(sdt)) > CDate(edt) Then
        outMsg = "INSUFFICIENT DATA - selected sample range is shorter than the " & h & "-year rolling horizon."
        ValidateInputs = False: Exit Function
    End If

    outMsg = ""
    ValidateInputs = True
End Function
