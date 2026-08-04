# vybe-test: powershell/arrays/array_slicing
$arr = @(10, 20, 30, 40, 50)
$slice = $arr[1..3]
if ($slice.Count -ne 3) {
    Write-Host "FAIL: expected 3 elements, got $($slice.Count)"
    exit 1
}
if ($slice[0] -ne 20) {
    Write-Host "FAIL: expected first element 20, got $($slice[0])"
    exit 1
}
if ($slice[2] -ne 40) {
    Write-Host "FAIL: expected last element 40, got $($slice[2])"
    exit 1
}
Write-Host "PASS"
exit 0
