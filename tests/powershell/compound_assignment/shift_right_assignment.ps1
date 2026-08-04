# vybe-test: powershell/compound_assignment/shift_right_assignment
$x = 2
$x >>= 1
if ($x -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
