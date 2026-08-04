# vybe-test: powershell/implicit_returns/hashtable_output
function Make { @{Name='x'} }
if ((Make).Name -eq 'x') { exit 0 }
exit 1
