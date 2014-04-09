Attribute VB_Name = "bond_formulas"
Option Explicit
'
' Built-in Excel Functions
'
' YIELD seems to return BEY
'
' DOLLARDE converts fractional price (1/32nds) to decimal price
'
' ACCRINT accrued interest...set issue date = date of prior coupon
' COUPPCD   date of prior coupon
' COUPNCD   date of next coupon
'
' COUPDAYBS number of days since beginning of coupon period
' COUPDAYS  number of days in this coupon period
'
' Accrued Interest = (COUPDAYBS/COUPDAYS) * coupon_rate/2 * face
' (1 - w)C    w = days_to_next/days_in_period ((COUPDAYS-COUPDAYBS)/COUPDAYS)
'
' PRICE returns clean price, given quoted (dirty) yield
' Invoice (dirty) price = PRICE + Accrued Interest = PV of cash flows
'
'

'
' Hull p77 Discrete/Continuous interest rate conversion
'
Function r_continuous(r_discrete As Double, m As Integer) As Double
    '
    ' convert r_discrete, which is compounded "m" times a year (1=annual, 2=sa)
    ' to continuous
    '
    r_continuous = m * Log(1 + r_discrete / m)
End Function
Function r_discrete(r_continuous As Double, m As Integer) As Double
    '
    ' convert r_continuous to r_discrete which is compounded "m" times a year
    '
    r_discrete = m * (Exp(r_continuous / m) - 1)
End Function

Function DV_01(mod_duration As Double, price As Double) As Double

    ' somehow, it won't let me define "DV01"
    
    DV_01 = (mod_duration * price) / 10000
End Function

Function bond_price(coupon_annual_percent As Double, _
                     term_years As Double, _
                     par_value As Double, _
                     annual_compounding_frequency As Integer, _
                     ytm_percent As Double)
                     
    Dim Periods As Double, m As Integer, T As Integer, _
        PV As Double, y As Double, n As Double, CF As Double, sum As Double, _
        numerator As Double, denominator As Double
    
    m = annual_compounding_frequency
    y = ytm_percent
    n = term_years
    Periods = m * n
    CF = par_value * coupon_annual_percent / m
    
    PV = 0
    For T = 1 To Periods
        PV = PV + CF / (1 + y / m) ^ T
        If T = Periods Then
            PV = PV + par_value / (1 + y / m) ^ T
        End If
    Next T
    
    bond_price = PV
                     
End Function
Function convexity(coupon_annual_percent As Double, _
                     term_years As Double, _
                     par_value As Double, _
                     annual_compounding_frequency As Integer, _
                     ytm_percent As Double)
                     
    Dim Periods As Double, m As Integer, T As Integer, _
        PV As Double, y As Double, n As Double, CF As Double, sum As Double, _
        numerator As Double, denominator As Double
    
    m = annual_compounding_frequency
    y = ytm_percent
    n = term_years
    Periods = m * n
    CF = par_value * coupon_annual_percent / m
    
    ' find the price (PV)
    PV = bond_price(coupon_annual_percent, n, par_value, m, y)
    
    ' find the convexity
    sum = 0
    For T = 1 To Periods
        numerator = T * (T + 1) * CF
        If T = Periods Then
            numerator = numerator + T * (T + 1) * par_value
        End If
        denominator = (1 + y / m) ^ (T + 2) * m ^ 2
        sum = sum + numerator / denominator
    Next T
    
    convexity = 1 / PV * sum
                     
End Function
                    

Function BEY(EAY As Double) As Double

    ' Bond Equivalent Yield  ... semi-annual EAY
    '
    '       1               1                 1
    '   ---------  =  -------------  =  -------------
    '   (1 + EAY)     (1 + BEY/2)^2     (1 + Y_m/m)^m
    '
    ' (1 + EAY)                 = (1 + BEY/2)^2
    ' (1 + EAY)^(1/2)           = 1 + BEY/2
    ' (1 + EAY)^(1/2) - 1       = BEY/2
    ' ((1 + EAY)^(1/2) - 1) * 2 = BEY
    '
    ' "m" = compounding periods per year

    BEY = ((1 + EAY) ^ (1 / 2) - 1) * 2
End Function

Function EAY(discrete_yield As Double, _
             compounding_periods_per_year As Integer) As Double
             
    ' Effective Annual Yield
    ' EAY = (1 + Y_m/m)^m -1
    
    Dim y As Double, m As Double
    y = discrete_yield
    m = compounding_periods_per_year
    
    EAY = ((1 + y / m) ^ m) - 1
             
End Function

' Macaulay Duration

Function MacDur(coupon_annual_percent As Double, _
                term_years As Double, _
                par_value As Double, _
                annual_compounding_frequency As Integer, _
                ytm_percent As Double)

    Dim Periods As Double, T As Integer, m As Integer, _
        PV As Double, y As Double, n As Double, CF As Double, sum As Double, _
        numerator As Double, denominator As Double
    
    m = annual_compounding_frequency
    y = ytm_percent
    n = term_years
    Periods = m * n
    CF = par_value * coupon_annual_percent / m
    
    ' find the price (PV)
    PV = bond_price(coupon_annual_percent, n, par_value, m, y)
    
    sum = 0
    For T = 1 To Periods
        sum = sum + CF / (1 + y / m) ^ T * T / m
        If T = Periods Then
            sum = sum + par_value / (1 + y / m) ^ T * T / m
        End If
    Next T
    
    MacDur = sum / PV

End Function
Function ModDur(coupon_annual_percent As Double, _
                term_years As Double, _
                par_value As Double, _
                annual_compounding_frequency As Integer, _
                ytm_percent As Double)
                
Dim y As Double, m As Integer
y = ytm_percent
m = annual_compounding_frequency

ModDur = MacDur(coupon_annual_percent, term_years, par_value, m, y) / _
         (1 + y / m)
End Function

' For Modified Duration use the Excel formula
' MDURATION(settlement, maturity, coupon, yld, frequency, [basis])
'
' basis
' 0 or omitted US (NASD) 30/360
' 1 Actual/Actual       US Treasuries (aka Act/Act ICMA ISMA -99 Act/Act ISMA)
' 2 Actual/360          commercial paper, T-bills (aka Act/360 A/360 French)
' 3 Actual/365          Treasury Bonds
' 4 European 30 / 360   corporate bonds, U.S. Agency bonds, and all mortgage backed securities

' ----------------------------------------------------------
' Based on DLucas' MIT Fixed Income lectures

Function coupon_total_return(annual_coupon_pct As Double, _
                             compounding_periods As Integer, _
                             face As Double, _
                             reinvestment_interest_rate As Double, _
                             term_in_years As Double) As Double
                             
    Dim c As Double, m As Integer, n As Double, y As Double
    c = annual_coupon_pct * face
    m = compounding_periods
    n = term_in_years
    y = reinvestment_interest_rate
    
    ' the FV of a fixed annuity
    coupon_total_return = c / m * (((1 + y / m) ^ (m * n)) - 1) / (y / m)
    
End Function

Function coupon_return(annual_coupon_pct As Double, _
                       compounding_periods As Integer, _
                       face As Double, _
                       term_in_years As Double) As Double
                       
    Dim c As Double, m As Integer, n As Double
    c = annual_coupon_pct
    m = compounding_periods
    n = term_in_years
    
    coupon_return = n * m * c / m * face
                             
End Function

Function interest_on_interest(annual_coupon_pct As Double, _
                             compounding_periods As Integer, _
                             face As Double, _
                             reinvestment_interest_rate As Double, _
                             term_in_years As Double) As Double
                             
    interest_on_interest = coupon_total_return(annual_coupon_pct, _
                                               compounding_periods, _
                                               face, _
                                               reinvestment_interest_rate, _
                                               term_in_years) _
                          - _
                          coupon_return(annual_coupon_pct, _
                                        compounding_periods, _
                                        face, _
                                        term_in_years)
                       

End Function

Function total_return(annual_coupon_pct As Double, _
                      compounding_periods As Integer, _
                      face As Double, _
                      reinvestment_interest_rate As Double, _
                      term_in_years As Double, _
                      purchase_price As Double, _
                      Optional redemption_value As Double) As Double
                      
    ' NOTE: Redemption Value is what you get when you dispose of the bond
    '       This is the face amount if the bond matures (default)
    '       It is the sale price if you sell it
        
    
    If IsMissing(redemption_value) = True Then
        redemption_value = face
    End If
                      
    total_return = coupon_total_return(annual_coupon_pct, _
                                       compounding_periods, _
                                       face, _
                                       reinvestment_interest_rate, _
                                       term_in_years) _
                    + _
                    redemption_value - purchase_price

End Function

Function expected_yield(annual_coupon_pct As Double, _
                        compounding_periods As Integer, _
                        face As Double, _
                        reinvestment_interest_rate As Double, _
                        term_in_years As Double, _
                        purchase_price As Double, _
                        Optional redemption_value As Double) As Double
                        
    ' NOTE: Redemption Value is what you get when you dispose of the bond
    '       This is the face amount if the bond matures (default)
    '       It is the sale price if you sell it
                      
    If IsMissing(redemption_value) = True Then
        redemption_value = face
    End If
                        
    Dim tot_return As Double
    tot_return = total_return(annual_coupon_pct, _
                              compounding_periods, _
                              face, _
                              reinvestment_interest_rate, _
                              term_in_years, _
                              purchase_price, _
                              redemption_value)
    
    Dim m As Integer, n As Double
    m = compounding_periods
    n = term_in_years
    
    expected_yield = m * (((tot_return + purchase_price) / purchase_price) ^ (1 / (m * n)) - 1)

End Function

