# vybe-test: powershell/data_conversion/array_to_string
$x = (1,2,3).ToString()
if ($x -ne $null) { exit 0 }
exit 1
