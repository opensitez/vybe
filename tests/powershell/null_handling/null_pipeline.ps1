# vybe-test: powershell/null_handling/null_pipeline
if (($null | Where-Object { $_ }) -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
