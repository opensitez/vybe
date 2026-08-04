# vybe-test: powershell/here_strings/empty_herestring
$here = @"
"@
if ($here -eq "\n") { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
