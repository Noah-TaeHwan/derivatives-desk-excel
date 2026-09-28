Attribute VB_Name = "DeskPricing"
Option Explicit

' Pricing and risk engine for the scenario P&L and arbitrage sheets.
' No add-ins, ActiveX, network access or auto-run events.
' Input rates are decimal fractions, e.g. 3.75% = 0.0375.
' Greeks are computed from the model, never hardcoded from reference values.
' Bond convention:
' compound full coupon periods and simple discount for the stub period.

Private Const DD_BP As Double = 0.0001
Private Const DD_MAX_PAYMENTS As Long = 600

Public Function DD_BondPrice(ByVal face As Double, ByVal coupon As Double, _
    ByVal ytm As Double, ByVal settle As Date, ByVal maturity As Date, _
    Optional ByVal freq As Long = 2) As Variant
    On Error GoTo Bad
    Dim rr As Double, w As Double, pv As Double, cf As Double
    Dim stepMonths As Long, n As Long, i As Long
    Dim dd As Date, nextCoupon As Date
    If face <= 0 Or coupon < 0 Or settle >= maturity Then GoTo Bad
    If freq <> 1 And freq <> 2 And freq <> 4 And freq <> 12 Then GoTo Bad
    If settle < DateSerial(1900, 3, 1) Then GoTo Bad
    If maturity > DateSerial(9998, 12, 31) Then GoTo Bad
    If settle <> DateValue(settle) Or maturity <> DateValue(maturity) Then GoTo Bad
    rr = ytm / freq
    If 1 + rr <= 0 Then GoTo Bad
    stepMonths = 12 / freq
    Do
        dd = DateSerial(Year(maturity), Month(maturity) - n * stepMonths, Day(maturity))
        If dd <= settle Then Exit Do
        n = n + 1
        If n > DD_MAX_PAYMENTS Then GoTo Bad
    Loop
    If n = 0 Then GoTo Bad
    nextCoupon = DateSerial(Year(maturity), _
        Month(maturity) - (n - 1) * stepMonths, Day(maturity))
    If nextCoupon <= dd Then GoTo Bad
    w = CDbl(nextCoupon - settle) / CDbl(nextCoupon - dd)
    If 1 + rr * w <= 0 Then GoTo Bad
    For i = 0 To n - 1
        cf = face * coupon / freq
        If i = n - 1 Then cf = cf + face
        pv = pv + cf / ((1 + rr) ^ i * (1 + rr * w))
    Next i
    DD_BondPrice = pv
    Exit Function
Bad:
    DD_BondPrice = CVErr(2015)
End Function

' Positive unit DV01/PV01 magnitude. Short position sign comes from quantity.
Public Function DD_BondPV01(ByVal face As Double, ByVal coupon As Double, _
    ByVal ytm As Double, ByVal settle As Date, ByVal maturity As Date, _
    Optional ByVal freq As Long = 2) As Variant
    On Error GoTo Bad
    Dim up As Variant, down As Variant
    down = DD_BondPrice(face, coupon, ytm - DD_BP, settle, maturity, freq)
    up = DD_BondPrice(face, coupon, ytm + DD_BP, settle, maturity, freq)
    If IsError(up) Or IsError(down) Then GoTo Bad
    DD_BondPV01 = (CDbl(down) - CDbl(up)) / 2
    Exit Function
Bad:
    DD_BondPV01 = CVErr(2015)
End Function

' Teaching convention: simple carry less a cash dividend at maturity.
Public Function DD_FuturePrice(ByVal spot As Double, ByVal rate As Double, _
    ByVal days As Double, ByVal cashDividend As Double, _
    Optional ByVal yearDays As Double = 365) As Variant
    On Error GoTo Bad
    If spot <= 0 Or days < 0 Or yearDays <= 0 Or cashDividend < 0 Then GoTo Bad
    If 1 + rate * days / yearDays <= 0 Then GoTo Bad
    DD_FuturePrice = spot * (1 + rate * days / yearDays) - cashDividend
    Exit Function
Bad:
    DD_FuturePrice = CVErr(2015)
End Function

Private Function DD_N(ByVal x As Double) As Double
    DD_N = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function DD_Np(ByVal x As Double) As Double
    DD_Np = Exp(-0.5 * x * x) / Sqr(8 * Atn(1))
End Function

' European Black-Scholes-Merton model with a continuous dividend yield.
' metric: PRICE, DELTA, GAMMA, VEGA, THETA
' VEGA is per 1.00 volatility; THETA is per calendar YEAR, not per day.
' Unit Greeks are unscaled; quantity/multiplier scaling happens on the sheets.
' Expired options and zero volatility intentionally return #VALUE!:
' these teaching formulas require strictly positive time and volatility.
Public Function DD_Option(ByVal metric As String, ByVal cp As String, _
    ByVal spot As Double, ByVal strike As Double, ByVal days As Double, _
    ByVal rate As Double, ByVal sigma As Double, ByVal divYield As Double, _
    Optional ByVal yearDays As Double = 365) As Variant
    On Error GoTo Bad
    Dim t As Double, d1 As Double, d2 As Double, sqt As Double
    Dim eq As Double, er As Double, dens As Double, ans As Double
    cp = LCase$(Trim$(cp))
    metric = UCase$(Trim$(metric))
    If cp <> "c" And cp <> "p" Then GoTo Bad
    If spot <= 0 Or strike <= 0 Or days <= 0 Then GoTo Bad
    If sigma <= 0 Or yearDays <= 0 Then GoTo Bad
    t = days / yearDays
    sqt = Sqr(t)
    d1 = (Log(spot / strike) + (rate - divYield + 0.5 * sigma ^ 2) * t) / (sigma * sqt)
    d2 = d1 - sigma * sqt
    eq = Exp(-divYield * t)
    er = Exp(-rate * t)
    dens = DD_Np(d1)
    Select Case metric
        Case "PRICE"
            If cp = "c" Then
                ans = spot * eq * DD_N(d1) - strike * er * DD_N(d2)
            Else
                ans = strike * er * DD_N(-d2) - spot * eq * DD_N(-d1)
            End If
        Case "DELTA"
            If cp = "c" Then
                ans = eq * DD_N(d1)
            Else
                ans = eq * (DD_N(d1) - 1)
            End If
        Case "GAMMA"
            ans = eq * dens / (spot * sigma * sqt)
        Case "VEGA"
            ans = spot * eq * dens * sqt
        Case "THETA"
            ans = -spot * eq * dens * sigma / (2 * sqt)
            If cp = "c" Then
                ans = ans - rate * strike * er * DD_N(d2) + divYield * spot * eq * DD_N(d1)
            Else
                ans = ans + rate * strike * er * DD_N(-d2) - divYield * spot * eq * DD_N(-d1)
            End If
        Case Else
            GoTo Bad
    End Select
    DD_Option = ans
    Exit Function
Bad:
    DD_Option = CVErr(2015)
End Function

Public Function DD_Position(ByVal unitRisk As Double, ByVal quantity As Double, _
    ByVal multiplier As Double) As Variant
    On Error GoTo Bad
    If multiplier <= 0 Then GoTo Bad
    DD_Position = unitRisk * quantity * multiplier
    Exit Function
Bad:
    DD_Position = CVErr(2015)
End Function

' shock is bp for BOND, a decimal stock/volatility change for other risks.
' referenceStep is 0.01 for DELTA/GAMMA/VEGA/FUTURE and 1 day for THETA.
Public Function DD_PnL(ByVal kind As String, ByVal positionRisk As Double, _
    ByVal shock As Double, Optional ByVal referenceStep As Double = 0.01) As Variant
    On Error GoTo Bad
    If referenceStep <= 0 Then GoTo Bad
    Select Case UCase$(Trim$(kind))
        Case "BOND"
            DD_PnL = -positionRisk * shock
        Case "GAMMA"
            DD_PnL = 0.5 * positionRisk * (shock / referenceStep) ^ 2
        Case "DELTA", "VEGA", "FUTURE", "THETA"
            DD_PnL = positionRisk * shock / referenceStep
        Case Else
            GoTo Bad
    End Select
    Exit Function
Bad:
    DD_PnL = CVErr(2015)
End Function


' Recalculate only this workbook.
Public Sub RecalculateAll()
    ThisWorkbook.Worksheets(1).Calculate
    MsgBox "VBA calculation complete.", vbInformation, "Scenario P&L"
End Sub

' Check live VBA functions and zero-position behaviour. Restore inputs on every exit.
Public Sub CheckZeroPosition()
    Dim sh As Worksheet, c As Range, positions As Variant
    Dim savedBond As Variant, savedOption As Variant, savedFuture As Variant
    Dim checks As Long, failure As String
    On Error GoTo Failed
    Set sh = ThisWorkbook.Worksheets(1)
    savedBond = sh.Range("G9").Value2
    savedOption = sh.Range("J12").Value2
    savedFuture = sh.Range("J13").Value2
    sh.Calculate
    For Each c In sh.Range("A1:L36")
        If IsError(c.Value) Then Err.Raise 5, , "Formula error at " & c.Address
    Next c
    checks = checks + 1
    If Abs(DD_PnL("GAMMA", 10, 0.02) - DD_PnL("GAMMA", 10, -0.02)) > 1E-07 Then Err.Raise 5, , "Gamma symmetry"
    checks = checks + 1
    If Not IsError(DD_Option("PRICE", "c", 0, 74000, 25, 0.0375, 0.1907, 0)) Then Err.Raise 5, , "Invalid spot guard"
    checks = checks + 1
    sh.Range("G9").Value2 = 0
    sh.Range("J12").Value2 = 0
    sh.Range("J13").Value2 = 0
    sh.Calculate
    For Each c In sh.Range("E23:K27")
        If IsError(c.Value) Then Err.Raise 5, , "Zero-position error at " & c.Address
        If Abs(CDbl(c.Value2)) > 1E-07 Then Err.Raise 5, , "Zero-position PnL at " & c.Address
    Next c
    checks = checks + 1
    GoTo RestoreInputs
Failed:
    failure = Err.Description
RestoreInputs:
    On Error Resume Next
    sh.Range("G9").Value2 = savedBond
    sh.Range("J12").Value2 = savedOption
    sh.Range("J13").Value2 = savedFuture
    sh.Calculate
    On Error GoTo 0
    If Len(failure) = 0 Then
        MsgBox CStr(checks) & "/4 VBA checks passed. Inputs restored.", vbInformation, "Scenario P&L"
    Else
        MsgBox "Check failed: " & failure & ". Inputs restored.", vbExclamation, "Scenario P&L"
    End If
End Sub
