# vybe-test: powershell/string_literal_quotes/escaped_single_quote
if ('It''s' -eq "It's") { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
