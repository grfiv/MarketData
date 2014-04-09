Attribute VB_Name = "GetQuoteYahoo_function"
Option Explicit

Function GetQuoteYahoo(Symbol As String, _
         Optional YahooParameter As String = "LastTradePriceOnly", _
         Optional YahooFeed As String = "a") As String

    ' This is an Excel VBA function which returns data about a stock
    ' using Yahoo Finance's various XML streams
    
    ' On December 8, 2012 =GetQuoteYahoo("VTI")                   returned 73.04
    '                     =GetQuoteYahoo("VTI", "Dividend_Yield") returned 1.89
    '                     =GetQuoteYahoo("MS",  "Industry", "e")  returned Investment Brokerage - National
    
    ' This function returns a string;
    ' the Str_to_Num function included below will convert a string-number to a double
    
    ' Yahoo Finance offers several XML streams
    ' You should look at each stream and figure out which one you want and what the node names are
    ' Stream description
    '   "a" is from the CSV data
    '   "b" is from yahoo.finance.quotes
    '   "c" and "d" are from yahoo.finance.quant and quant2
    '   "e" is from yahoo.finance.stocks
    '
    ' =GetQuoteYahoo("GOOG","URL","a") will give you the URL for stream "a" to look at for Google, etc.
    ' The default is stream "a" and node "LastTradePriceOnly"
    
    ' The full list (as of Dec 2012) of yahoo.finance data tables was
    '   yahoo.finance.historicaldata
    '   yahoo.finance.industry
    '   yahoo.finance.isin
    '   yahoo.finance.onvista
    '   yahoo.finance.option_contracts
    '   yahoo.finance.options
    '   yahoo.finance.quant
    '   yahoo.finance.quant2
    '   yahoo.finance.quotes
    '   yahoo.finance.quoteslist
    '   yahoo.finance.sectors
    '   yahoo.finance.stock
    '   yahoo.finance.stocks
    '   yahoo.finance.xchange
    
    ' In case you want even more variety, I have a GetQuote function which uses the XML
    ' stream from Google Finance but they are threatening to discontinue it
    ' see http://www.philadelphia-reflections.com/blog/2385.htm
    
    ' My thanks to http://vikku.info/codetrash/Yahoo_Finance_Stock_Quote_API
    ' and http://developer.yahoo.com/yql/console/
    
    ' Before you start, read this:
    ' http://stackoverflow.com/questions/11245733/declaring-early-bound-msxml-object-throws-an-error-in-vba
    
    ' --------- code follows ---------
    
    Dim YahooXMLstream As MSXML2.DOMDocument
    
    Dim oChildren As MSXML2.IXMLDOMNodeList
    Dim oChild As MSXML2.IXMLDOMNode
    
    Dim fSuccess As Boolean
    Dim URL As String, _
        url_part1 As String, _
        url_part2 As String, _
        url_part3 As String, _
        url_part4 As String, _
        url_part5 As String
    
    On Error GoTo HandleErr
    
    ' create the URL that requests the XML stream from Yahoo Finance
    If LCase(YahooFeed) = "a" Then
        url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20csv%20where%20url%3D'http%3A%2F%2Fdownload.finance.yahoo.com%2Fd%2Fquotes.csv%3Fs%3D"
        url_part2 = "%26f%3Dsnll1d1t1cc1p2t7va2ibb6aa5pomwj5j6k4k5ers7r1qdyj1t8e7e8e9r6r7r5b4p6p5j4m3m7m8m4m5m6k1b3b2i5x"
        url_part3 = "%26e%3D.csv'%20and%20columns%3D"
        url_part4 = "'Symbol%2CName%2CLastTradeWithTime%2CLastTradePriceOnly%2CLastTradeDate%2CLastTradeTime%2CChange%20PercentChange%2CChange%2CChangeinPercent%2CTickerTrend%2CVolume%2CAverageDailyVolume%2CMoreInfo%2CBid%2CBidSize%2CAsk%2CAskSize%2CPreviousClose%2COpen%2CDayRange%2CFiftyTwoWeekRange%2CChangeFromFiftyTwoWeekLow%2CPercentChangeFromFiftyTwoWeekLow%2CChangeFromFiftyTwoWeekHigh%2CPercentChangeFromFiftyTwoWeekHigh%2CEarningsPerShare%2CPE%20Ratio%2CShortRatio%2CDividendPayDate%2CExDividendDate%2CDividendPerShare%2CDividend%20Yield%2CMarketCapitalization%2COneYearTargetPrice%2CEPS%20Est%20Current%20Yr%2CEPS%20Est%20Next%20Year%2CEPS%20Est%20Next%20Quarter%2CPrice%20per%20EPS%20Est%20Current%20Yr%2CPrice%20per%20EPS%20Est%20Next%20Yr%2CPEG%20Ratio%2CBook%20Value%2CPrice%20to%20Book%2CPrice%20to%20Sales%2CEBITDA"
        url_part5 = "%2CFiftyDayMovingAverage%2CChangeFromFiftyDayMovingAverage%2CPercentChangeFromFiftyDayMovingAverage%2CTwoHundredDayMovingAverage%2CChangeFromTwoHundredDayMovingAverage%2CPercentChangeFromTwoHundredDayMovingAverage%2CLastTrade%20(Real-time)%20with%20Time%2CBid%20(Real-time)%2CAsk%20(Real-time)%2COrderBook%20(Real-time)%2CStockExchange'"
    
        URL = url_part1 & Trim(Symbol) & url_part2 & url_part3 & url_part4 & url_part5
    ElseIf LCase(YahooFeed) = "b" Then
        url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20%28%22"
        url_part2 = "%22%29&diagnostics=false&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
        
        URL = url_part1 & Trim(Symbol) & url_part2
    ElseIf LCase(YahooFeed) = "c" Then
        url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quant%20where%20symbol%20in%20(%22"
        url_part2 = "%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
        
        URL = url_part1 & Trim(Symbol) & url_part2
    ElseIf LCase(YahooFeed) = "d" Then
        url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quant2%20where%20symbol%20in%20(%22"
        url_part2 = "%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
        
        URL = url_part1 & Trim(Symbol) & url_part2
    ElseIf LCase(YahooFeed) = "e" Then
        url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.stocks%20where%20symbol%20in%20(%22"
        url_part2 = "%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
        
        URL = url_part1 & Trim(Symbol) & url_part2
    Else
        ' return error message if YahooFeed isn't recognized
        GetQuoteYahoo = YahooFeed & " is an invalid YahooFeed parameter supplied to GetQuoteYahoo function"
        Exit Function
    End If
    
    ' In case you want to look at the XML
    If YahooParameter = "URL" Then
        GetQuoteYahoo = URL
        Exit Function
    End If
    
    ' pull in the XML stream
    Set YahooXMLstream = New MSXML2.DOMDocument
    YahooXMLstream.async = False                 ' wait for completion
    YahooXMLstream.validateOnParse = False       ' do not validate the XML stream
    
    fSuccess = YahooXMLstream.Load(URL)          ' pull in the feed
    
    If Not fSuccess Then                         ' quit on failure
      MsgBox "error loading Yahoo Finance XML stream"
      Exit Function
    End If
    
    ' iterate through the nodes looking for one with the name in YahooParameter
    GetQuoteYahoo = YahooParameter & " is not valid for GetQuoteYahoo function with YahooFeed " & YahooFeed
    
    Set oChildren = YahooXMLstream.DocumentElement.LastChild.LastChild.ChildNodes
    
    For Each oChild In oChildren

        If oChild.nodeName = YahooParameter Then
        
            GetQuoteYahoo = oChild.Text
            Exit Function
            
        End If
        
    Next oChild
        
' error handlers
ExitHere:
            Exit Function
HandleErr:
            MsgBox "Error " & Err.Number & ": " & Err.Description
            Resume ExitHere
            Resume
End Function

Function Str_to_Num(num_as_string As String) As Double

    ' The GetQuoteYahoo function returns string values
    ' This function converts numbers in string format to double

    Str_to_Num = Val(num_as_string)

End Function
