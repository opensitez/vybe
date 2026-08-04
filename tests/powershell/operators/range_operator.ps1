# vybe-test: powershell/operators/range_operator
$range = 1..5
$count = $range.Count
if ($count -ne 5) {
    Write-Host "FAIL: expected 5, got $count"
    exit 1
}
if ($range[0] -ne 1) {
    Write-Host "FAIL: expected first element 1, got $($range[0])"
    exit 1
}
if ($range[4] -ne 5) {
    Write-Host "FAIL: expected last element 5, got $($range[4])"
    exit 1
}
Write-Host "PASS"
exit 0
