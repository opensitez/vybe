# vybe-test: powershell/arrays/array_sort
$arr = @(3, 1, 2)
$sorted = $arr | Sort-Object
if ($sorted[0] -ne 1) {
    Write-Host "FAIL: expected first element 1, got $($sorted[0])"
    exit 1
}
if ($sorted[2] -ne 3) {
    Write-Host "FAIL: expected last element 3, got $($sorted[2])"
    exit 1
}
Write-Host "PASS"
exit 0
