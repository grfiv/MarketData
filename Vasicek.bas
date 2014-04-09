Attribute VB_Name = "Vasicek"
Option Explicit
'
'  Vasicek One-factor short-rate term-structure model
'
'  regress r_t+1 = a + b * r_t + c * epsilon
'
'  c = stdev(residuals) [sigma * sqrt(delta) and v * sqrt(h)]
'
'               Jun Pan                                    Hui Chen
'   ---------------------------------------     ------------------------------
'    kappa                                      gamma
'    delta          (1/12 for monthly data)     h
'    sigma = stdev(residuals) / sqrt(delta)     v = stdev(residuals) / sqrt(h)
'
'
Function Vasicek_kappa(b, delta)                ' aka gamma(b, h)
    Vasicek_kappa = (1 - b) / delta
End Function

Function Vasicek_rbar(a, kappa, delta)          ' aka rbar(a, gamma, h)
    Vasicek_rbar = a / (kappa * delta)
End Function

Function Vasicek_sigma(stdev_residuals, delta)  ' aka v(stdev_residuals, h)
    Vasicek_sigma = stdev_residuals / Sqr(delta)
End Function
' take stdev of Vasicek_variance for more-accurate sigma (v)
Function Vasicek_variance(variance_residuals, kappa, delta)
    Vasicek_variance = variance_residuals * 2 * kappa / (1 - Exp(-2 * kappa * delta))
End Function

Function Vasicek_alpha(rbar, kappa, T, sigma)   ' aka A(rbar, gamma, T, v)
    Dim PartA, PartA1 As Double
    PartA1 = (1 - Exp(-kappa * T)) / kappa
    PartA = rbar * (PartA1 - T)
    
    Dim PartB1, PartB2, PartB3, PartB As Double
    PartB1 = (sigma * sigma) / (2 * kappa * kappa)
    PartB2 = (1 - Exp(-2 * kappa * T)) / (2 * kappa)
    PartB3 = 2 * ((1 - Exp(-kappa * T)) / kappa)
    PartB = PartB1 * (PartB2 - PartB3 + T)
    
    Vasicek_alpha = PartA + PartB
End Function

Function Vasicek_beta(kappa, term)              ' aka B(gamma, term)
    Vasicek_beta = (Exp(-kappa * term) - 1) / kappa
End Function

Function Vasicek_price(alpha, beta, r)          ' aka price(A, B, r)
    Vasicek_price = Exp(alpha + beta * r)
End Function

Function Vasicek_yield(price, term)
    Vasicek_yield = Log(1 / price) / term
End Function

