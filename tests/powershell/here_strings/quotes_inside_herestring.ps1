# vybe-test: powershell/here_strings/quotes_inside_herestring
$here = @"
She said "Hello"
"@
if ($here -match 'She said') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
