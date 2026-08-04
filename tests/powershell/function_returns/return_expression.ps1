# vybe-test: powershell/function_returns/return_expression
function Eval { return 2+3 }
if ((Eval) -eq 5) { exit 0 }
exit 1
