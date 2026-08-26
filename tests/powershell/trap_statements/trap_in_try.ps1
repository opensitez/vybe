# vybe-test: powershell/trap_statements/trap_in_try
$caught = $false
trap { $script:caught = $true; continue }
try { throw 'ERR' } finally { }
if ($caught) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
