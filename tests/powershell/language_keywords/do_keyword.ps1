# vybe-test: powershell/language_keywords/do_keyword
$i = 1
do { Write-Host 'PASS'; exit 0 } while ($false)
Write-Host 'FAIL'
exit 1
