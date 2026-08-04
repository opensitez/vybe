# vybe-test: powershell/here_strings/simple_herestring_single
$here = @'
Hello
World
'@
if ($here -match 'Hello') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
