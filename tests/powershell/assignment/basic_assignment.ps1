# vybe-test: powershell/assignment/basic_assignment
$x = 5
if ($x -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
