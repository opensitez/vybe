# vybe-test: powershell/data_conversion/string_to_int
$x = [int]'6'
if ($x -eq 6) { exit 0 }
exit 1
