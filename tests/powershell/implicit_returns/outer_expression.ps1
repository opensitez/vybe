# vybe-test: powershell/implicit_returns/outer_expression
function Outer { if ($true) { 'yes' } }
if ((Outer) -eq 'yes') { exit 0 }
exit 1
