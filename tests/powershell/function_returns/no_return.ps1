# vybe-test: powershell/function_returns/no_return
function Get-Nothing { 'hi' }
if ((Get-Nothing) -eq 'hi') { exit 0 }
exit 1
