# vybe-test: powershell/literals/array_literals
$arr = 1,2,3
if ($arr.Length -ne 3 -or $arr[0] -ne 1) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
