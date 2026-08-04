# vybe-test: powershell/implicit_returns/if_expression
function Check { param($x) if ($x) { 'yes' } }
if ((Check $true) -eq 'yes') { exit 0 }
exit 1
