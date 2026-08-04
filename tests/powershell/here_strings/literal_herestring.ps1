# vybe-test: powershell/here_strings/literal_herestring
$here = @'
$value
'@
if ($here -match '\$value') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
