Attribute VB_Name = "GetQuote_function"
Option Explicit

Function GetQuote(Symbol As String, Optional GoogleParameter As String = "last") As String

    ' This is an Excel VBA function which returns data about a stock
    ' using Google Finance's XML stream
    
    ' On December 6, 2012 the Excel call =GetQuote("VTI")           returned 72.55
    '                                    =GetQuote("VTI", "volume") returned 68676
    
    ' Since this function returns strings, I've included a String_to_Number function below
    
    ' I got the essential ideas from here:
    ' http://msdn.microsoft.com/en-us/library/aa163921%28office.10%29.aspx
    
    ' BUT before you start, read this:
    ' http://stackoverflow.com/questions/11245733/declaring-early-bound-msxml-object-throws-an-error-in-vba
    
    ' "GoogleParameter" is one of the following node names, it defaults to "last":
    '
    '<xml_api_reply version="1">
    '    <finance module_id="0" tab_id="0" mobile_row="0" mobile_zipped="1" row="0" section="0">
    '        <symbol data="VTI"/>
    '        <pretty_symbol data="VTI"/>
    '        <symbol_lookup_url data="/finance?client=ig&q=VTI"/>
    '        <company data="Vanguard Total Stock Market ETF"/>
    '        <exchange data="NYSEARCA"/>
    '        <exchange_timezone data=""/>
    '        <exchange_utc_offset data=""/>
    '        <exchange_closing data=""/>
    '        <divisor data="2"/>
    '        <currency data="USD"/>
    '        <last data="72.55"/>
    '        <high data="72.69"/>
    '        <low data="72.45"/>
    '        <volume data="68676"/>
    '        <avg_volume data=""/>
    '        <market_cap data="22385.76"/>
    '        <open data="72.54"/>
    '        <y_close data="72.60"/>
    '        <change data="+0.03"/>
    '        <perc_change data="0.04"/>
    '        <delay data="0"/>
    '        <trade_timestamp data="1 minute ago"/>
    '        <trade_date_utc data="20121206"/>
    '        <trade_time_utc data="145144"/>
    '        <current_date_utc data="20121206"/>
    '        <current_time_utc data="145323"/>
    '        <symbol_url data="/finance?client=ig&q=VTI"/>
    '        <chart_url data="/finance/chart?q=NYSEARCA:VTI&tlf=12"/>
    '        <disclaimer_url data="/help/stock_disclaimer.html"/>
    '        <ecn_url data=""/>
    '        <isld_last data="72.55"/>
    '        <isld_trade_date_utc data="20121206"/>
    '        <isld_trade_time_utc data="142903"/>
    '        <brut_last data=""/>
    '        <brut_trade_date_utc data=""/>
    '        <brut_trade_time_utc data=""/>
    '        <daylight_savings data="false"/>
    '    </finance>
    '</xml_api_reply>
    
    Dim GoogleXMLstream As MSXML2.DOMDocument
    
    Dim oChildren As MSXML2.IXMLDOMNodeList
    Dim oChild As MSXML2.IXMLDOMNode
    
    Dim fSuccess As Boolean
    Dim URL As String
    
    On Error GoTo HandleErr
    
    ' create the URL that requests the XML stream from Google Finance
    URL = "http://www.google.com/ig/api?stock=" & Trim(Symbol)
    
    ' pull in the XML stream
    Set GoogleXMLstream = New MSXML2.DOMDocument
    GoogleXMLstream.async = False                 ' wait for completion
    GoogleXMLstream.validateOnParse = False       ' do not validate the XML stream
    
    fSuccess = GoogleXMLstream.Load(URL)
    
    If Not fSuccess Then                          ' quit on failure
      MsgBox "error loading Google Finance XML stream"
      Exit Function
    End If
    
    ' iterate through the nodes looking for one with the name in GoogleParameter
    GetQuote = GoogleParameter & " is not valid for GetQuote function"
    
    Set oChildren = GoogleXMLstream.DocumentElement.LastChild.ChildNodes
    
    For Each oChild In oChildren
    
        If oChild.nodeName = GoogleParameter Then
        
            GetQuote = oChild.Attributes.getNamedItem("data").Text
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

Function String_to_Number(num_as_string As String) As Double

    ' The GetQuote function returns string values
    ' This function converts numbers in string format to double

    String_to_Number = Val(num_as_string)

End Function

