# vybe-test: powershell/cmdlets/select_object_first
$numbers = 1..10
$first3 = $numbers | Select-Object -First 3
if ($first3.Count -ne 3) {
    Write-Host "FAIL: expected 3 elements, got $($first3.Count)"
    exit 1
}
if ($first3[2] -ne 3) {
    Write-Host "FAIL: expected last element 3, got $($first3[2])"
    exit 1
}
Write-Host "PASS"
exit 0
