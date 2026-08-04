# vybe-test: powershell/operators/array_multiplication
$arr = @(1, 2) * 3
if ($arr.Count -ne 6) {
    Write-Host "FAIL: expected 6 elements, got $($arr.Count)"
    exit 1
}
if ($arr[0] -ne 1 -or $arr[2] -ne 1 -or $arr[4] -ne 1) {
    Write-Host "FAIL: array multiplication pattern incorrect"
    exit 1
}
Write-Host "PASS"
exit 0
