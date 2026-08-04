# vybe-test: powershell/data_conversion/bool_to_string
$x = $true
if ($x.ToString() -eq 'True') { exit 0 }
exit 1
