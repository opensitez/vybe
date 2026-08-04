# vybe-test: powershell/implicit_returns/array_output
function GetValues { 1,2,3 }
if ((GetValues).Count -eq 3) { exit 0 }
exit 1
