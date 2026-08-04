# vybe-test: powershell/format_strings/expand_string
$name = 'World'
if ("Hello $name" -eq 'Hello World') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
