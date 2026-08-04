# vybe-test: powershell/here_strings/multiline_herestring
$here = @"
Line1
Line2
Line3
"@
if ($here -match 'Line2') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
