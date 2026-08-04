# vybe-test: powershell/property_access/command_result_chain
if ((New-Object System.Text.StringBuilder).Length -eq 0) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
