# vybe-test: powershell/arithmetic_operators/increment
$x = 1
$x += 1
if ($x -ne 2) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
