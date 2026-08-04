# vybe-test: powershell/compound_assignment/multiply_assignment
$x = 2
$x *= 3
if ($x -ne 6) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
