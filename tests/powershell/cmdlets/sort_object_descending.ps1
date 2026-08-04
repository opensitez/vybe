# vybe-test: powershell/cmdlets/sort_object_descending
$numbers = @(5, 2, 8, 1, 9)
$sorted = $numbers | Sort-Object -Descending
if ($sorted[0] -ne 9) {
    Write-Host "FAIL: expected first element 9, got $($sorted[0])"
    exit 1
}
if ($sorted[4] -ne 1) {
    Write-Host "FAIL: expected last element 1, got $($sorted[4])"
    exit 1
}
Write-Host "PASS"
exit 0
