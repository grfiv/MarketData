Attribute VB_Name = "GetHistoricalData_function"
Option Explicit

Function GetHistoricalData(Symbol As String, _
                           QuoteDate As Date, _
                  Optional QuoteType As String = "AdjClose") As Double

    ' Returns stock data for "Symbol" on "QuoteDate" using Yahoo Finance
    '
    ' The choices for "QuoteType" are
    '   Open
    '   High
    '   Low
    '   Close
    '   Volume
    '   Adj Close or AdjClose (Default)
    '
    ' ... and these calculated values:
    '   MAX (maximum of Open, High, Low, Close, AdjClose)
    '   MIN (minimum of Open, High, Low, Close, AdjClose)
    '   AVG (average of High, Low)
    '
    ' for example
    '    =GetHistoricalData("BRK.A", DATEVALUE("2/26/2012"))
    ' returns
    '    120,350.00
    '
    ' (you'd be more likely to refer to a cell with a date in it)
    '
    ' Thanks to Peter Urbani at http://www.wilmott.com/messageview.cfm?catid=10&threadid=25730
    '
    ' Note: I figure out if you gave me a weekend and I look for the previous Friday
    '       but if you give me a weekday holiday, I will produce unpredictable results
    '       give me "02/30/1998" and I'll give you #VALUE
    
    
    ' If you want current data see the following:
    '
    ' http://www.philadelphia-reflections.com/blog/2392.htm
    ' http://www.philadelphia-reflections.com/blog/2385.htm
    
    ' Before you start, read this:
    ' http://stackoverflow.com/questions/11245733/declaring-early-bound-msxml-object-throws-an-error-in-vba
    
    Dim URL As String
    
    Dim StartMonth As Integer, _
        EndMonth As Integer, _
        StartDay As Integer, _
        EndDay As Integer, _
        StartYear As Integer, _
        EndYear As Integer, _
        DateInt As Integer
    
    Dim Parts() As String

    ' if date entered is a weekend, find the previous Friday
    DateInt = Weekday(QuoteDate)
    If (DateInt = 1) Then      ' Sunday
        QuoteDate = DateAdd("d", -2, QuoteDate)
    ElseIf (DateInt = 7) Then  ' Saturday
        QuoteDate = DateAdd("d", -1, QuoteDate)
    End If

    ' note that I pick a single date
    StartYear = year(QuoteDate)
    EndYear = StartYear

    StartMonth = month(QuoteDate)
    EndMonth = StartMonth

    StartDay = day(QuoteDate)
    EndDay = StartDay

    ' Yahoo Finance URL
    URL = "http://ichart.finance.yahoo.com/table.csv?s=" & Symbol & _
           IIf(StartMonth = 0, "&a=0", "&a=" & (StartMonth - 1)) & _
           IIf(StartDay = 0, "&b=1", "&b=" & StartDay) & _
           IIf(StartYear = 0, "&c=" & EndYear, "&c=" & StartYear) & _
           IIf(EndMonth = 0, "", "&d=" & (EndMonth - 1)) & _
           IIf(EndDay = 0, "", "&e=" & EndDay) & _
           IIf(EndYear = 0, "", "&f=" & EndYear) & _
           "&g=d" & _
           "&ignore=.csv"


    ' Send the request URL
    Dim HTTP As New XMLHTTP
    HTTP.Open "GET", URL, False
    HTTP.Send
    
    If HTTP.Status <> "200" Then
        MsgBox "request error: " & HTTP.Status
        Exit Function
    End If
    
    ' split the returned comma-delimited string at the commas
    Parts = Split(HTTP.responseText, ",")
    
    Select Case LCase(QuoteType)
        Case "open"
            GetHistoricalData = Val(Parts(7))
            Exit Function
        Case "high"
            GetHistoricalData = Val(Parts(8))
            Exit Function
        Case "low"
            GetHistoricalData = Val(Parts(9))
            Exit Function
        Case "close"
            GetHistoricalData = Val(Parts(10))
            Exit Function
        Case "volume"
            GetHistoricalData = Val(Parts(11))
            Exit Function
        Case "adjclose", "adj close"
            GetHistoricalData = Val(Parts(12))
            Exit Function
        Case "max"
            GetHistoricalData = Application.Max(Val(Parts(7)), Val(Parts(8)), Val(Parts(9)), Val(Parts(10)), Val(Parts(12)))
            Exit Function
        Case "min"
            GetHistoricalData = Application.Min(Val(Parts(7)), Val(Parts(8)), Val(Parts(9)), Val(Parts(10)), Val(Parts(12)))
            Exit Function
        Case "avg"
            GetHistoricalData = Application.Average(Val(Parts(8)), Val(Parts(9)))
            Exit Function
        Case Else
            MsgBox QuoteType & " invalid QuoteType for GetHistoricalData function"
            Exit Function
    End Select
    
End Function
