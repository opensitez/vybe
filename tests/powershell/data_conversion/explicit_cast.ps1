# vybe-test: powershell/data_conversion/explicit_cast
$x = [string](5)
if ($x -eq '5') { exit 0 }
exit 1
