# vybe-test: powershell/function_returns/return_from_switch
function Match { param($x) switch ($x) { 1 { return 'one' } default { return 'other' } } }
if ((Match 1) -eq 'one' -and (Match 5) -eq 'other') { exit 0 }
exit 1
