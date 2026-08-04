# vybe-test: powershell/command_results/captured_error
$value = $null
try { $null.Method() } catch { $value = 'PASS' }
if ($value -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
