# vybe-test: powershell/dynamic_typing/boolean_conversion
$x = 0
if (-not $x) { exit 0 }
exit 1
