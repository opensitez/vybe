# vybe-test: powershell/information_streams/information_pipeline
1..3 | ForEach-Object { Write-Information 'i' }
Write-Host 'PASS'
exit 0
