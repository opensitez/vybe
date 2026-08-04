# vybe-test: powershell/data_conversion/implicit_conversion
$x = 5 + '5'
if ($x -eq '55') { exit 0 }
exit 1
