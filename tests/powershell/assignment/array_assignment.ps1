# vybe-test: powershell/assignment/array_assignment
$arr = 1,2,3
if ($arr[0] -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
