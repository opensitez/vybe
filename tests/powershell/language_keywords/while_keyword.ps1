# vybe-test: powershell/language_keywords/while_keyword
$i = 1
while ($i -eq 1) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
