# vybe-test: powershell/trap_statements/trap_with_error_action
$caught = $false
trap { $caught = $true; continue }
Get-ChildItem no_such -ErrorAction Stop
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
