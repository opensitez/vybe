# vybe-test: powershell/data_conversion/string_to_bool
$x = [bool]'True'
if ($x -eq $true) { exit 0 }
exit 1
