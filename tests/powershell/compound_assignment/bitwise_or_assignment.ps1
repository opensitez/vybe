# vybe-test: powershell/compound_assignment/bitwise_or_assignment
$x = 1
$x |= 2
if ($x -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
