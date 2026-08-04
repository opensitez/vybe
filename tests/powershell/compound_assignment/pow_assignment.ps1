# vybe-test: powershell/compound_assignment/pow_assignment
$x = 2
$x **= 3
if ($x -ne 8) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
