# vybe-test: powershell/function_returns/multiple_return
function Choose { param($x) if ($x) { return 'yes' } return 'no' }
if ((Choose $true) -eq 'yes' -and (Choose $false) -eq 'no') { exit 0 }
exit 1
