Attribute VB_Name = "Module2"
Function BondPV01(face As Double, couponRate As Double, ytm As Double, settle As Date, maturity As Date, Optional freq As Long = 2) As Double
    Dim pDown As Double
    Dim pUp As Double
    pDown = BondDirtyPrice(face, couponRate, ytm - 0.0001, settle, maturity, freq)
    pUp = BondDirtyPrice(face, couponRate, ytm + 0.0001, settle, maturity, freq)
    BondPV01 = (pDown - pUp) / 2
End Function
