# vybe-test: powershell/language_keywords/continue_keyword
$i = 0
while ($i -lt 1) { $i = $i + 1; continue; Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
