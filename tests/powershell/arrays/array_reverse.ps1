# vybe-test: powershell/arrays/array_reverse
$arr = @(1, 2, 3)
[array]::Reverse($arr)
if ($arr[0] -ne 3) {
    Write-Host "FAIL: expected first element 3, got $($arr[0])"
    exit 1
}
if ($arr[2] -ne 1) {
    Write-Host "FAIL: expected last element 1, got $($arr[2])"
    exit 1
}
Write-Host "PASS"
exit 0
