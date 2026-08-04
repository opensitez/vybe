# vybe-test: powershell/compound_assignment/bitwise_xor_assignment
$x = 3
$x ^= 1
if ($x -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
