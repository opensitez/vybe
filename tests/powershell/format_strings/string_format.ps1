# vybe-test: powershell/format_strings/string_format
if ([string]::Format('{0} {1}','Hello','World') -eq 'Hello World') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
