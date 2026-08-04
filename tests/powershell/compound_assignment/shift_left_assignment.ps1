# vybe-test: powershell/compound_assignment/shift_left_assignment
$x = 1
$x <<= 1
if ($x -ne 2) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
