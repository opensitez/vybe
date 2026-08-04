# vybe-test: powershell/compound_assignment/divide_assignment
$x = 6
$x /= 2
if ($x -ne 3) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
