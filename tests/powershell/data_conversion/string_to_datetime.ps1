# vybe-test: powershell/data_conversion/string_to_datetime
$x = [datetime]'2020-01-01'
if ($x.Year -eq 2020) { exit 0 }
exit 1
