# vybe-test: powershell/data_conversion/int_to_string
$x = 5
if ($x.ToString() -eq '5') { exit 0 }
exit 1
