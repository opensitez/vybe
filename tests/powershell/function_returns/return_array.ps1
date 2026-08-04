# vybe-test: powershell/function_returns/return_array
function Get-Array { return ,(1,2,3) }
if ((Get-Array).Count -eq 3) { exit 0 }
exit 1
