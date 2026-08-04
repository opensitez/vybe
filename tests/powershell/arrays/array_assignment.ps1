# vybe-test: powershell/arrays/array_assignment
$arr = @(1, 2, 3)
$arr[1] = 99
if ($arr[1] -ne 99) {
    Write-Host "FAIL: expected 99, got $($arr[1])"
    exit 1
}
Write-Host "PASS"
exit 0
