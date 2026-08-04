# vybe-test: powershell/command_quoting/escape_in_single_quotes
if ((Write-Output 'It''s OK') -eq "It's OK") { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
