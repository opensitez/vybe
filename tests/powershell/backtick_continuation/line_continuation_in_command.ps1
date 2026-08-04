# vybe-test: powershell/backtick_continuation/line_continuation_in_command
$result = Get-Process `
| Where-Object { $_.Id -gt 0 }
Write-Host 'PASS'
exit 0
