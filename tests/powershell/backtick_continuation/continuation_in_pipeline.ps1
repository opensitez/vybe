# vybe-test: powershell/backtick_continuation/continuation_in_pipeline
Get-Process `
| Select-Object -First 1
Write-Host 'PASS'
exit 0
