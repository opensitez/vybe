# vybe-test: powershell/cmdlets/select_object_last
$numbers = 1..10
$last2 = $numbers | Select-Object -Last 2
if ($last2.Count -ne 2) {
    Write-Host "FAIL: expected 2 elements, got $($last2.Count)"
    exit 1
}
if ($last2[1] -ne 10) {
    Write-Host "FAIL: expected last element 10, got $($last2[1])"
    exit 1
}
Write-Host "PASS"
exit 0
