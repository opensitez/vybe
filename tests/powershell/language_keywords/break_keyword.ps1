# vybe-test: powershell/language_keywords/break_keyword
for ($i = 0; $i -lt 1; $i++) { break; Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
