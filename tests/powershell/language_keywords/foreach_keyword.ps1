# vybe-test: powershell/language_keywords/foreach_keyword
foreach ($i in 1..1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
