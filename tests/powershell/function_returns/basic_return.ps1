# vybe-test: powershell/function_returns/basic_return
function Get-Value { return 5 }
if ((Get-Value) -eq 5) { exit 0 }
exit 1
