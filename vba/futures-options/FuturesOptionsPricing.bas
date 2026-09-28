Attribute VB_Name = "Module3"
' Pricing module - Futures + Black-Scholes + Greeks
' Import: File > Import File (keeps Attribute line). Paste manually: delete Attribute line first.
Option Explicit

Private Function NCDF(x As Double) As Double
    NCDF = Application.WorksheetFunction.Norm_S_Dist(x, True)
End Function

Private Function NPDF(x As Double) As Double
    NPDF = Application.WorksheetFunction.Norm_S_Dist(x, False)
End Function

Private Function CalcD1(S As Double, K As Double, T As Double, r As Double, sigma As Double, q As Double) As Double
    If T <= 0 Or sigma <= 0 Or S <= 0 Or K <= 0 Then
        CalcD1 = 0
    Else
        CalcD1 = (Log(S / K) + (r - q + 0.5 * sigma * sigma) * T) / (sigma * Sqr(T))
    End If
End Function

Private Function CalcD2(d1 As Double, sigma As Double, T As Double) As Double
    CalcD2 = d1 - sigma * Sqr(T)
End Function

' Futures: S=spot, r=rate, Tdays=days, DivAmt=dividend amount
' Ex: =FuturesPrice(73400,0.0375,27,0) -> 73604
Public Function FuturesPrice(S As Double, r As Double, Tdays As Double, DivAmt As Double) As Double
    Dim T As Double
    If Tdays < 0 Then Tdays = 0
    T = Tdays / 365
    FuturesPrice = S * (1 + r * T) - DivAmt
End Function

' Black-Scholes price. CP="C" call, "P" put
' Ex: =BSPrice(72600,74000,25,0.0375,0.1907,0,"C") -> about 931.5
Public Function BSPrice(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String) As Double
    Dim T As Double, d1 As Double, d2 As Double
    T = Tdays / 365
    If T <= 0 Then
        If UCase(CP) = "C" Then
            BSPrice = Application.WorksheetFunction.Max(S - K, 0)
        Else
            BSPrice = Application.WorksheetFunction.Max(K - S, 0)
        End If
        Exit Function
    End If
    d1 = CalcD1(S, K, T, r, sigma, q)
    d2 = CalcD2(d1, sigma, T)
    If UCase(CP) = "C" Then
        BSPrice = S * Exp(-q * T) * NCDF(d1) - K * Exp(-r * T) * NCDF(d2)
    Else
        BSPrice = K * Exp(-r * T) * NCDF(-d2) - S * Exp(-q * T) * NCDF(-d1)
    End If
End Function

Public Function BSDelta(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String) As Double
    Dim T As Double, d1 As Double
    T = Tdays / 365
    If T <= 0 Then
        If UCase(CP) = "C" Then
            BSDelta = IIf(S > K, 1, 0)
        Else
            BSDelta = IIf(S < K, -1, 0)
        End If
        Exit Function
    End If
    d1 = CalcD1(S, K, T, r, sigma, q)
    If UCase(CP) = "C" Then
        BSDelta = Exp(-q * T) * NCDF(d1)
    Else
        BSDelta = Exp(-q * T) * (NCDF(d1) - 1)
    End If
End Function

Public Function BSGamma(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double) As Double
    Dim T As Double, d1 As Double
    T = Tdays / 365
    If T <= 0 Or sigma <= 0 Then
        BSGamma = 0
        Exit Function
    End If
    d1 = CalcD1(S, K, T, r, sigma, q)
    BSGamma = Exp(-q * T) * NPDF(d1) / (S * sigma * Sqr(T))
End Function

' Vega per 1pct vol move
Public Function BSVega(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double) As Double
    Dim T As Double, d1 As Double
    T = Tdays / 365
    If T <= 0 Then
        BSVega = 0
        Exit Function
    End If
    d1 = CalcD1(S, K, T, r, sigma, q)
    BSVega = S * Exp(-q * T) * NPDF(d1) * Sqr(T) / 100
End Function

' Theta per 1 day (annual/252 trading days to match example)
Public Function BSTheta(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String) As Double
    Dim T As Double, d1 As Double, d2 As Double, annual As Double
    T = Tdays / 365
    If T <= 0 Then
        BSTheta = 0
        Exit Function
    End If
    d1 = CalcD1(S, K, T, r, sigma, q)
    d2 = CalcD2(d1, sigma, T)
    If UCase(CP) = "C" Then
        annual = (-S * sigma * Exp(-q * T) * NPDF(d1) / (2 * Sqr(T)) - r * K * Exp(-r * T) * NCDF(d2) + q * S * Exp(-q * T) * NCDF(d1))
    Else
        annual = (-S * sigma * Exp(-q * T) * NPDF(d1) / (2 * Sqr(T)) + r * K * Exp(-r * T) * NCDF(-d2) - q * S * Exp(-q * T) * NCDF(-d1))
    End If
    BSTheta = annual / 252
End Function

Public Function BSTotalPremium(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String, Qty As Double) As Double
    BSTotalPremium = BSPrice(S, K, Tdays, r, sigma, q, CP) * Qty
End Function

Public Function BSDeltaAmt(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String, Qty As Double) As Double
    BSDeltaAmt = BSDelta(S, K, Tdays, r, sigma, q, CP) * S * 0.01 * Qty
End Function

Public Function BSGammaAmt(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, Qty As Double) As Double
    BSGammaAmt = BSGamma(S, K, Tdays, r, sigma, q) * (S * 0.01) ^ 2 * Qty
End Function

Public Function BSVegaAmt(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, Qty As Double) As Double
    BSVegaAmt = BSVega(S, K, Tdays, r, sigma, q) * Qty
End Function

Public Function BSThetaAmt(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String, Qty As Double) As Double
    BSThetaAmt = BSTheta(S, K, Tdays, r, sigma, q, CP) * Qty
End Function

' Rho per 1pct rate move (bonus, not required)
Public Function BSRho(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String) As Double
    Dim T As Double, d2 As Double, d1 As Double
    T = Tdays / 365
    If T <= 0 Then
        BSRho = 0
        Exit Function
    End If
    d1 = CalcD1(S, K, T, r, sigma, q)
    d2 = CalcD2(d1, sigma, T)
    If UCase(CP) = "C" Then
        BSRho = K * T * Exp(-r * T) * NCDF(d2) / 100
    Else
        BSRho = -K * T * Exp(-r * T) * NCDF(-d2) / 100
    End If
End Function

Public Function BSRhoAmt(S As Double, K As Double, Tdays As Double, r As Double, sigma As Double, q As Double, CP As String, Qty As Double) As Double
    BSRhoAmt = BSRho(S, K, Tdays, r, sigma, q, CP) * Qty
End Function

