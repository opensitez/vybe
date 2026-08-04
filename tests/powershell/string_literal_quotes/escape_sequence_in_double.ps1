# vybe-test: powershell/string_literal_quotes/escape_sequence_in_double
if ("Line`nEnd" -match 'End') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
