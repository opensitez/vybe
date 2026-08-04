# vybe-test: powershell/string_methods/contains
if ('PowerShell'.Contains('Shell')) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
