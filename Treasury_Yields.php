<?php
//
// Returns an array of the latest (usually the day before) yields on constant maturity Treasury bonds
//     see http://www.treasury.gov/resource-center/data-chart-center/interest-rates/Pages/TextView.aspx?data=yield
//
// Thanks very much to uramihsayibok, gmail, com at http://php.net/manual/en/function.simplexml-load-file.php

/*
On Sunday December 3, 2012 the following call:

    $yields = TreasuryYields();
    print_r($yields);

produced this output:

Array
(
    [Id] => 5736
    [NEW_DATE] => 2012-11-30T00:00:00
    [BC_1MONTH] => 0.11
    [BC_3MONTH] => 0.08
    [BC_6MONTH] => 0.13
    [BC_1YEAR] => 0.18
    [BC_2YEAR] => 0.25
    [BC_3YEAR] => 0.34
    [BC_5YEAR] => 0.61
    [BC_7YEAR] => 1.04
    [BC_10YEAR] => 1.62
    [BC_20YEAR] => 2.37
    [BC_30YEAR] => 2.81
    [BC_30YEARDISPLAY] => 2.81
)
*/

function TreasuryYields()
{
  // create the XML URL for today's month and year

  //    this one picks the last (more than) 5,000 entries (25+ years?)
  //    $url = 'http://data.treasury.gov/feed.svc/DailyTreasuryYieldCurveRateData';
  $url_part1      = 'http://data.treasury.gov/feed.svc/DailyTreasuryYieldCurveRateData?$filter=month(NEW_DATE)%20eq%20';
  $url_this_month = date("m");
  $url_part2      = '%20and%20year(NEW_DATE)%20eq%20';
  $url_this_year  = date("Y");

  $url = $url_part1 . $url_this_month . $url_part2 . $url_this_year;
  $xml = simplexml_load_file($url);
  $number_of_entries = count($xml->entry);

  // test for no entries (first day of the month on a Saturday, for example, has no entries for this month)
  if ($number_of_entries == 0)
  {
    $last_month     = date('Y-m-d', strtotime(date('Y-m-d')." -1 month"));
    $url_this_month = substr($last_month, 5, 2);
    $url_this_year  = substr($last_month, 0, 4);

    $url = $url_part1 . $url_this_month . $url_part2 . $url_this_year;
    $xml = simplexml_load_file($url);
    $number_of_entries = count($xml->entry);
  }

  // pull out all the yields for the last (most-recent) entry
  $counter = 0;
  foreach ($xml->entry as $entry)
  {
    if (++$counter < $number_of_entries) continue;
    foreach ($entry->content->children("m", true)->properties->children("d", true) as $label => $value)
    {
      $yields[$label] = (string)$value;
    }
  }
  return $yields;
}
?>