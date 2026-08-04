# vybe-test: powershell/here_strings/command_output_herestring
$here = @"
$(Write-Output 'PASS')
"@
if ($here -match 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
