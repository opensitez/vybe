# vybe-test: powershell/compound_assignment/bitwise_and_assignment
$x = 3
$x &= 1
if ($x -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
