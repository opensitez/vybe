# vybe-test: powershell/assignment/multiline_assignment
$x = 
    10
if ($x -ne 10) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
