# vybe-test: powershell/information_streams/information_with_object
Write-Information ([PSCustomObject]@{ A = 1 })
Write-Host 'PASS'
exit 0
