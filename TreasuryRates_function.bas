Attribute VB_Name = "TreasuryRates_function"
Option Explicit

Function TreasuryRates(Optional Maturity As String = "BC_10YEAR") As String

    ' This Excel VBA function returns a string with
    ' the most-recent constant-maturity Treasury bond yield specified
    
    ' The default is the 10-year rate
    
    ' Before you start, read this:
    ' http://stackoverflow.com/questions/11245733/declaring-early-bound-msxml-object-throws-an-error-in-vba
    
    ' The Maturity input argument is any one of the elements
    ' following the colon in the <d: ... nodes shown in a sample of the XML produced by the Treasury
    
    ' On December 4, 2012
    ' =TreasuryRates()            returned 1.62
    ' =TreasuryRates("BC_30YEAR") returned 2.78
    
    '  <entry>
    '    <content type="application/xml">
    '      <m:properties>
    '        <d:Id m:type="Edm.Int32">5738</d:Id>
    '        <d:NEW_DATE m:type="Edm.DateTime">2012-12-04T00:00:00</d:NEW_DATE>
    '        <d:BC_1MONTH m:type="Edm.Double">0.07</d:BC_1MONTH>
    '        <d:BC_3MONTH m:type="Edm.Double">0.1</d:BC_3MONTH>
    '        <d:BC_6MONTH m:type="Edm.Double">0.15</d:BC_6MONTH>
    '        <d:BC_1YEAR m:type="Edm.Double">0.18</d:BC_1YEAR>
    '        <d:BC_2YEAR m:type="Edm.Double">0.25</d:BC_2YEAR>
    '        <d:BC_3YEAR m:type="Edm.Double">0.34</d:BC_3YEAR>
    '        <d:BC_5YEAR m:type="Edm.Double">0.63</d:BC_5YEAR>
    '        <d:BC_7YEAR m:type="Edm.Double">1.04</d:BC_7YEAR>
    '        <d:BC_10YEAR m:type="Edm.Double">1.62</d:BC_10YEAR>
    '        <d:BC_20YEAR m:type="Edm.Double">2.36</d:BC_20YEAR>
    '        <d:BC_30YEAR m:type="Edm.Double">2.78</d:BC_30YEAR>
    '        <d:BC_30YEARDISPLAY m:type="Edm.Double">2.78</d:BC_30YEARDISPLAY>
    '      </m:properties>
    '    </content>
    '  </entry>
    
    ' This function returns a string; the String_to_Number function
    ' defined with my GetQuote function will convert to a string number to double
    
    Dim TreasuryXMLstream As MSXML2.DOMDocument
    
    Dim DNodes As MSXML2.IXMLDOMNodeList
    Dim DNode As MSXML2.IXMLDOMNode
    
    Dim fSuccess As Boolean
    Dim URL As String, _
        url_part1 As String, _
        url_this_month As String, _
        url_part2 As String, _
        url_this_year As String
    Dim iInt As Integer
    
    On Error GoTo HandleErr
    
    ' create the XML request URL for today's month and year
    ' -----------------------------------------------------
    
    ' this one returns more than 5,700 entries (25+ years?)
    ' URL = "http://data.treasury.gov/feed.svc/DailyTreasuryYieldCurveRateData"
    
    url_part1 = "http://data.treasury.gov/feed.svc/DailyTreasuryYieldCurveRateData?$filter=month(NEW_DATE)%20eq%20"
    url_part2 = "%20and%20year(NEW_DATE)%20eq%20"
    
    iInt = Month(Now)
    url_this_month = LTrim(Str(iInt))
    
    iInt = Year(Now)
    url_this_year = LTrim(Str(iInt))
    
    ' this is used to test whether the month we requested had no data
    ' ---------------------------------------------------------------
    TreasuryRates = "empty"
    
    ' set up to pull in the XML stream
    ' --------------------------------
    
    Set TreasuryXMLstream = New MSXML2.DOMDocument
    TreasuryXMLstream.async = False                 ' wait for completion
    TreasuryXMLstream.validateOnParse = False       ' do not validate the XML stream
    
TryAgain:
    
    URL = url_part1 & url_this_month & url_part2 & url_this_year
    
    ' pull in the XML
    ' ---------------
    
    fSuccess = TreasuryXMLstream.Load(URL)         ' load the XML stream
    
    If Not fSuccess Then                           ' quit on failure
        MsgBox "error loading Treasury XML stream"
        Exit Function
    End If
    
    ' Iterate through the <d: nodes looking for the <d:Maturity
    ' ---------------------------------------------------------
    
    ' this assumes
    ' 1. the last node in the XML stream returned to us is the <entry> node we want
    '   2. the last node in the <entry> node is a <content node
    '     3. the last node in the <content node is an <m:properties> node ...
    '       4. ... which contains the <d:BC_10YEAR (or whatever) nodes

    Set DNodes = TreasuryXMLstream.DocumentElement.LastChild.LastChild.LastChild.ChildNodes
    '                                              entry     content   m         d's
    
    For Each DNode In DNodes
    
        If DNode.BaseName = Maturity Then
                        
            TreasuryRates = DNode.Text
            Exit Function
                            
        End If
    
    Next DNode
    
    ' test for no entries (first day of the month on a Saturday, for example, has no entries for this month)
    ' go to prior month in that case
    
    If TreasuryRates = "empty" Then
    
        ' go through twice, and we assume the input parameter is wrong
        TreasuryRates = Maturity & " is not a valid parameter for TreasuryRates function"
        
        url_this_month = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "mm")
        url_this_year = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "yyyy")
        
        GoTo TryAgain
        
    End If
    
' error handlers
' --------------
    
ExitHere:
            Exit Function
HandleErr:
            MsgBox "Error " & Err.Number & ": " & Err.Description
            Resume ExitHere
            Resume
End Function
