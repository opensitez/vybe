# vybe-test: powershell/arrays/array_add_operator
$arr = @(1, 2)
$arr = $arr + 3
$count = $arr.Count
if ($count -ne 3) {
    Write-Host "FAIL: expected 3, got $count"
    exit 1
}
if ($arr[2] -ne 3) {
    Write-Host "FAIL: expected arr[2] = 3, got $($arr[2])"
    exit 1
}
Write-Host "PASS"
exit 0
