# vybe-test: powershell/language_keywords/switch_keyword
switch (2) { 2 { Write-Host 'PASS'; exit 0 } }
Write-Host 'FAIL'
exit 1
