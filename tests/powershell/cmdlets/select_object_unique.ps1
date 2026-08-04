# vybe-test: powershell/cmdlets/select_object_unique
$numbers = @(1, 2, 2, 3, 3, 3, 4)
$unique = $numbers | Select-Object -Unique
if ($unique.Count -ne 4) {
    Write-Host "FAIL: expected 4 unique elements, got $($unique.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
