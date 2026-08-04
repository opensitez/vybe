# vybe-test: powershell/arithmetic_operators/decrement
$x = 2
$x -= 1
if ($x -ne 1) { Write-Host 'FAIL'; exit 1 }
Write-Host 'PASS'
exit 0
