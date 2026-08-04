# vybe-test: powershell/function_returns/return_from_if
function Check { param($x) if ($x -eq 1) { return 'one' } 'none' }
if ((Check 1) -eq 'one' -and (Check 2) -eq 'none') { exit 0 }
exit 1
