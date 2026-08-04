# vybe-test: powershell/here_strings/subexpression_herestring
$here = @"
$(1 + 1)
"@
if ($here -match '2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
