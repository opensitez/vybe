# vybe-test: powershell/here_strings/simple_herestring_double
$here = @"
PowerShell
Test
"@
if ($here -match 'PowerShell') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
