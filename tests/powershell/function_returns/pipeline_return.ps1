# vybe-test: powershell/function_returns/pipeline_return
function Emit { 1; 2; 3 }
if ((Emit | Measure-Object).Count -eq 3) { exit 0 }
exit 1
