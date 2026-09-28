Attribute VB_Name = "Module1"
Function BondDirtyPrice(face As Double, couponRate As Double, ytm As Double, settle As Date, maturity As Date, Optional freq As Long = 2) As Double
    Dim coup As Double
    Dim r As Double
    Dim mStep As Long
    Dim prevCpn As Date
    Dim nextCpn As Date
    Dim dt As Date
    Dim nPer As Long
    Dim i As Long
    Dim w As Double
    Dim periodDays As Double
    Dim pv As Double
    Dim cf As Double

    coup = face * couponRate / freq
    r = ytm / freq
    mStep = 12 / freq

    dt = maturity
    nPer = 0
    Do While dt > settle
        nPer = nPer + 1
        dt = DateAdd("m", -mStep, dt)
    Loop
    prevCpn = dt
    nextCpn = DateAdd("m", mStep, prevCpn)

    periodDays = CDbl(nextCpn - prevCpn)
    w = CDbl(nextCpn - settle) / periodDays

    pv = 0
    For i = 0 To nPer - 1
        If i = nPer - 1 Then
            cf = coup + face
        Else
            cf = coup
        End If
        pv = pv + cf / (1 + r) ^ (w + i)
    Next i

    BondDirtyPrice = pv
End Function
