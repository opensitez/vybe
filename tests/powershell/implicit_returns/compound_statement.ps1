# vybe-test: powershell/implicit_returns/compound_statement
function Show { 'a'; 'b' }
if ((Show | Select-Object -First 1) -eq 'a') { exit 0 }
exit 1
