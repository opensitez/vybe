# vybe-test: powershell/variable_promotion/function_block
function SetX { param($value) $x = $value }
SetX 6
if ($x -eq $null) { exit 0 }
exit 1
