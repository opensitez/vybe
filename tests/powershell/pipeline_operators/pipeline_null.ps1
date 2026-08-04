# vybe-test: powershell/pipeline_operators/pipeline_null
if (($null | ForEach-Object { $_ }) -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
