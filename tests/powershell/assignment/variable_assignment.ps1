# vybe-test: powershell/assignment/variable_assignment
$name = 'X'
if ($name -ne 'X') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
