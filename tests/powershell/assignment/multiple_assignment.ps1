# vybe-test: powershell/assignment/multiple_assignment
$x = $y = 1
if ($x -ne 1 -or $y -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
