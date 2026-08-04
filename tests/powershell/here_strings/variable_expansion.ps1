# vybe-test: powershell/here_strings/variable_expansion
$value = 'PASS'
$here = @"
$value
"@
if ($here -match 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
