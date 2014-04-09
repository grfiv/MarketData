<?php
//
// Takes in an array of symbols and returns an array of stock/fund information
// retrieved from various Yahoo Finance XML streams
//
// Thanks to http://vikku.info/codetrash/Yahoo_Finance_Stock_Quote_API
// and http://developer.yahoo.com/yql/console/
//
// The first parameter is an array of symbols
//
// The second parameter specifies the XML stream you want
//    "a"  CSV
//    "b"  yahoo.finance.quotes
//    "c"  yahoo.finance.quant
//    "d"  yahoo.finance.quant2
//    "e"  yahoo.finance.stocks
// or
//    "url_a", "url_b", etc.  to see the URL
// or
//    "all" returns an array with the results of all of the XML streams
//    This is the best way to figure out which stream and which nodes you want

/*    example

$symbols    = array("MS", "FNJHX");
$quote_list = YahooStockQuote($symbols, "all");
print_r($quote_list);

*/

function YahooStockQuote($symbols, $YahooStream = "a")
{
  // create the symbol list
  // ----------------------

  $symbol_list = '';
  foreach ($symbols as $symbol)
  {
    $symbol_list .= trim($symbol) . "%2C";                       # symbols separated by commas
  }
  $symbol_list   = substr($symbol_list, 0, -3);                  # strip off the last comma
  $symbol_list_r = str_replace('%2C','%22%2C%22',$symbol_list);  # different format for different streams

  // pull the specified XML stream and parse
  // ---------------------------------------

  if (strtolower(trim($YahooStream)) == "a" || strtolower(trim($YahooStream)) == "all" || strtolower(trim($YahooStream)) == "url_a")
  {
    $url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20csv%20where%20url%3D'http%3A%2F%2Fdownload.finance.yahoo.com%2Fd%2Fquotes.csv%3Fs%3D";
    $url_part2 = "%26f%3Dsnll1d1t1cc1p2t7va2ibb6aa5pomwj5j6k4k5ers7r1qdyj1t8e7e8e9r6r7r5b4p6p5j4m3m7m8m4m5m6k1b3b2i5x";
    $url_part3 = "%26e%3D.csv'%20and%20columns%3D";
    $url_part4 = "'Symbol%2CName%2CLastTradeWithTime%2CLastTradePriceOnly%2CLastTradeDate%2CLastTradeTime%2CChange%20PercentChange%2CChange%2CChangeinPercent%2CTickerTrend%2CVolume%2CAverageDailyVolume%2CMoreInfo%2CBid%2CBidSize%2CAsk%2CAskSize%2CPreviousClose%2COpen%2CDayRange%2CFiftyTwoWeekRange%2CChangeFromFiftyTwoWeekLow%2CPercentChangeFromFiftyTwoWeekLow%2CChangeFromFiftyTwoWeekHigh%2CPercentChangeFromFiftyTwoWeekHigh%2CEarningsPerShare%2CPE%20Ratio%2CShortRatio%2CDividendPayDate%2CExDividendDate%2CDividendPerShare%2CDividend%20Yield%2CMarketCapitalization%2COneYearTargetPrice%2CEPS%20Est%20Current%20Yr%2CEPS%20Est%20Next%20Year%2CEPS%20Est%20Next%20Quarter%2CPrice%20per%20EPS%20Est%20Current%20Yr%2CPrice%20per%20EPS%20Est%20Next%20Yr%2CPEG%20Ratio%2CBook%20Value%2CPrice%20to%20Book%2CPrice%20to%20Sales%2CEBITDA";
    $url_part5 = "%2CFiftyDayMovingAverage%2CChangeFromFiftyDayMovingAverage%2CPercentChangeFromFiftyDayMovingAverage%2CTwoHundredDayMovingAverage%2CChangeFromTwoHundredDayMovingAverage%2CPercentChangeFromTwoHundredDayMovingAverage%2CLastTrade%20(Real-time)%20with%20Time%2CBid%20(Real-time)%2CAsk%20(Real-time)%2COrderBook%20(Real-time)%2CStockExchange'";

    $URL = $url_part1 . $symbol_list . $url_part2 . $url_part3 . $url_part4 . $url_part5;

    if (strtolower(trim($YahooStream)) == "url_a")
    {
      echo $URL;
      exit;
    }

    // create a SimpleXML object from the XML stream
    $xml = simplexml_load_file($URL);

    // pull out the quotes
    foreach ($xml->results->row as $quote)
    {
        $symbol = (string)$quote->Symbol;
        foreach ($quote as $info_name => $info)
        {
        if ($info_name == "Symbol") continue;
        $node_list[$symbol][$info_name] = (string)$info;
        }
    }
    if (strtolower(trim($YahooStream)) == "a") return $node_list;

    $node_list_list["a"] = $node_list;
    unset($node_list);
  }

  if (strtolower(trim($YahooStream)) == "b" || strtolower(trim($YahooStream)) == "all" || strtolower(trim($YahooStream)) == "url_b")
  {
    $url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20%28%22";
    $url_part2 = "%22%29&diagnostics=false&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys";

    $URL = $url_part1 . $symbol_list . $url_part2;

    if (strtolower(trim($YahooStream)) == "url_b")
    {
      echo $URL;
      exit;
    }

    $xml = simplexml_load_file($URL);
    foreach ($xml->results->quote as $quote)
    {
        $symbol = (string)$quote->attributes()->symbol;
        foreach ($quote as $label => $value)
        {
          $node_list[$symbol][$label] = (string)$value;
        }
    }
    if (strtolower(trim($YahooStream)) == "b") return $node_list;

    $node_list_list["b"] = $node_list;
    unset($node_list);
  }

  if (strtolower(trim($YahooStream)) == "c" || strtolower(trim($YahooStream)) == "all" || strtolower(trim($YahooStream)) == "url_c")
  {
    $url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quant%20where%20symbol%20in%20(%22";
    $url_part2 = "%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys";

    $URL = $url_part1 . $symbol_list_r . $url_part2;

    if (strtolower(trim($YahooStream)) == "url_c")
    {
      echo $URL;
      exit;
    }

    $xml = simplexml_load_file($URL);
    foreach ($xml->results->stock as $quote)
    {
        $symbol = (string)$quote->attributes()->symbol;
        foreach ($quote as $label => $value)
        {
          $node_list[$symbol][$label] = (string)$value;
        }
    }
    if (strtolower(trim($YahooStream)) == "c") return $node_list;

    $node_list_list["c"] = $node_list;
    unset($node_list);
  }

  if (strtolower(trim($YahooStream)) == "d" || strtolower(trim($YahooStream)) == "all" || strtolower(trim($YahooStream)) == "url_d")
  {
    $url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quant2%20where%20symbol%20in%20(%22";
    $url_part2 = "%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys";

    $URL = $url_part1 . $symbol_list_r . $url_part2;

    if (strtolower(trim($YahooStream)) == "url_d")
    {
      echo $URL;
      exit;
    }

    $xml = simplexml_load_file($URL);
    foreach ($xml->results->stock as $quote)
    {
        $symbol = (string)$quote->attributes()->symbol;
        foreach ($quote as $label => $value)
        {
          $node_list[$symbol][$label] = (string)$value;
        }
    }
    if (strtolower(trim($YahooStream)) == "d") return $node_list;

    $node_list_list["d"] = $node_list;
    unset($node_list);
  }

  if (strtolower(trim($YahooStream)) == "e" || strtolower(trim($YahooStream)) == "all" || strtolower(trim($YahooStream)) == "url_e")
  {
    $url_part1 = "http://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.stocks%20where%20symbol%20in%20(%22";
    $url_part2 = "%22)&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys";

    $URL = $url_part1 . $symbol_list_r . $url_part2;

    if (strtolower(trim($YahooStream)) == "url_e")
    {
      echo $URL;
      exit;
    }

    $xml = simplexml_load_file($URL);
    foreach ($xml->results->stock as $quote)
    {
        $symbol = (string)$quote->attributes()->symbol;
        foreach ($quote as $label => $value)
        {
          $node_list[$symbol][$label] = (string)$value;
        }
    }
    if (strtolower(trim($YahooStream)) == "e") return $node_list;

    $node_list_list["e"] = $node_list;
    unset($node_list);

    return $node_list_list;
  }
}
?>