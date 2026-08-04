# vybe-test: powershell/data_conversion/date_to_string
$x = Get-Date -Date '2020-01-01'
if ($x.ToString().Contains('2020')) { exit 0 }
exit 1
