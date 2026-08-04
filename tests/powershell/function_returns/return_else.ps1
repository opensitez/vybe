# vybe-test: powershell/function_returns/return_else
function Decide { param($flag) if ($flag) { return 'yes' } else { return 'no' } }
if ((Decide $true) -eq 'yes' -and (Decide $false) -eq 'no') { exit 0 }
exit 1
