# vybe-test: powershell/implicit_returns/from_function
function Make { 'implicit' }
if ((Make) -eq 'implicit') { exit 0 }
exit 1
