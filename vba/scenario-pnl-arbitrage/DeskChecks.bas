Attribute VB_Name = "DeskChecks"
Option Explicit

' The shock is bp for bonds, days for theta, and a decimal fraction otherwise.
' Validate each caller's unit before using the original compiled pricing engine.
Public Function DD_PnL_Checked(ByVal kind As String, ByVal positionRisk As Double, _
    ByVal shock As Double, Optional ByVal referenceStep As Double = 0.01) As Variant
    On Error GoTo Bad
    Select Case UCase$(Trim$(kind))
        Case "BOND", "THETA"
            If Abs(referenceStep - 1#) > 1E-12 Then GoTo Bad
        Case "DELTA", "GAMMA", "VEGA", "FUTURE"
            If Abs(referenceStep - 0.01) > 1E-12 Then GoTo Bad
        Case Else
            GoTo Bad
    End Select
    DD_PnL_Checked = DD_PnL(kind, positionRisk, shock, referenceStep)
    Exit Function
Bad:
    DD_PnL_Checked = CVErr(2015)
End Function

Public Sub CheckScenarioPnL()
    Dim sh As Worksheet, c As Range
    On Error GoTo Failed
    Set sh = Sheet1
    sh.Calculate
    For Each c In sh.Range("A1:L36")
        If IsError(c.Value) Then Err.Raise 5, , "Formula error at " & c.Address
    Next c
    If Abs(DD_PnL_Checked("BOND", 10, 1, 1) + 10) > 1E-07 Then Err.Raise 5, , "Bond bp unit"
    If Abs(DD_PnL_Checked("DELTA", 10, 0.01, 0.01) - 10) > 1E-07 Then Err.Raise 5, , "Delta unit"
    If Abs(DD_PnL_Checked("GAMMA", 10, 0.01, 0.01) - 5) > 1E-07 Then Err.Raise 5, , "Gamma half factor"
    If Abs(DD_PnL_Checked("GAMMA", 10, 0.02, 0.01) - DD_PnL_Checked("GAMMA", 10, -0.02, 0.01)) > 1E-07 Then Err.Raise 5, , "Gamma symmetry"
    If Abs(DD_PnL_Checked("VEGA", 10, 0.01, 0.01) - 10) > 1E-07 Then Err.Raise 5, , "Vega unit"
    If Abs(DD_PnL_Checked("THETA", 10, 1, 1) - 10) > 1E-07 Then Err.Raise 5, , "Theta unit"
    If Abs(DD_PnL_Checked("FUTURE", 10, -0.01, 0.01) + 10) > 1E-07 Then Err.Raise 5, , "Future unit"
    If Not IsError(DD_PnL_Checked("DELTA", 10, 0.01, 1)) Then Err.Raise 5, , "Wrong delta step accepted"
    If Not IsError(DD_PnL_Checked("BOND", 10, 1, 0.01)) Then Err.Raise 5, , "Wrong bond step accepted"
    If Abs(CDbl(DD_Option("PRICE", "c", 72600, 74000, 25, 0.0375, 0.1907, 0, 365)) - 931.529621972877) > 1E-06 Then Err.Raise 5, , "Option baseline"
    If Abs(CDbl(DD_FuturePrice(72600, 0.0375, 25, 0, 365)) - 72786.4726027397) > 1E-06 Then Err.Raise 5, , "Future baseline"
    MsgBox "Scenario P&L audit passed: formulas, units, model baselines.", vbInformation, "Scenario P&L"
    Exit Sub
Failed:
    MsgBox "Scenario P&L audit failed: " & Err.Description, vbExclamation, "Scenario P&L"
End Sub

Public Sub CheckArbitrage()
    Dim sh As Worksheet, c As Range, rowNum As Long
    Dim fairBuy As Double, fairSell As Double
    On Error GoTo Failed
    Set sh = Sheet2
    sh.Calculate
    For Each c In sh.Range("A1:L251")
        If IsError(c.Value) Then Err.Raise 5, , "Formula error at " & c.Address
    Next c
    If Abs(CDbl(DD_FuturePrice(71300, 0.0375, 21, 0, 365)) - 71453.8321917808) > 1E-06 Then Err.Raise 5, , "Samsung future baseline"
    If Abs(CDbl(sh.Range("K9").Value2) - (CDbl(sh.Range("C9").Value2) - CDbl(sh.Range("I9").Value2))) > 1E-06 Then Err.Raise 5, , "Samsung gap"
    For rowNum = 15 To 19
        fairBuy = CDbl(DD_FuturePrice(sh.Cells(rowNum, 2).Value2, sh.Range("H5").Value2, sh.Range("G9").Value2, sh.Cells(rowNum, 6).Value2, sh.Range("K5").Value2))
        fairSell = CDbl(DD_FuturePrice(sh.Cells(rowNum, 3).Value2, sh.Range("H5").Value2, sh.Range("G9").Value2, sh.Cells(rowNum, 6).Value2, sh.Range("K5").Value2))
        If Abs(CDbl(sh.Cells(rowNum, 9).Value2) - (CDbl(sh.Cells(rowNum, 5).Value2) - fairBuy)) > 0.01 Then Err.Raise 5, , "Sell spread row " & rowNum
        If Abs(CDbl(sh.Cells(rowNum, 10).Value2) - (fairSell - CDbl(sh.Cells(rowNum, 4).Value2))) > 0.01 Then Err.Raise 5, , "Buy spread row " & rowNum
    Next rowNum
    If Abs(CDbl(sh.Range("G27").Value2) - CDbl(DD_FuturePrice(sh.Range("A27").Value2, sh.Range("H24").Value2, sh.Range("E27").Value2, sh.Range("K24").Value2, sh.Range("K29").Value2))) > 1E-08 Then Err.Raise 5, , "K200 future price"
    If Abs(CDbl(sh.Range("I32").Value2) + CDbl(sh.Range("C29").Value2) * CDbl(sh.Range("G29").Value2) * CDbl(sh.Range("I27").Value2)) > 0.01 Then Err.Raise 5, , "K200 gross gap"
    If CDbl(sh.Range("A39").Value2) <> 199 Then Err.Raise 5, , "Basket count"
    If Abs(CDbl(sh.Range("G39").Value2) - 1#) > 1E-10 Then Err.Raise 5, , "Basket weight"
    If Abs(CDbl(sh.Range("I244").Value2) - (CDbl(sh.Range("E32").Value2) - CDbl(sh.Range("L241").Value2))) > 0.01 Then Err.Raise 5, , "Basket rounding balance"
    MsgBox "Arbitrage audit passed: prices, spreads, index and basket.", vbInformation, "Arbitrage"
    Exit Sub
Failed:
    MsgBox "Arbitrage audit failed: " & Err.Description, vbExclamation, "Arbitrage"
End Sub
