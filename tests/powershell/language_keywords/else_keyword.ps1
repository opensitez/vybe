# vybe-test: powershell/language_keywords/else_keyword
if ($false) { } else { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
