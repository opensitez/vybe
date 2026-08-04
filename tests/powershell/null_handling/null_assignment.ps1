# vybe-test: powershell/null_handling/null_assignment
$x = $null
if ($x -eq $null) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
