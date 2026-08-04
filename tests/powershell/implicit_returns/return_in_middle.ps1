# vybe-test: powershell/implicit_returns/return_in_middle
function Example { 'first'; return 'stop'; 'later' }
if ((Example | Select-Object -First 1) -eq 'first') { exit 0 }
exit 1
