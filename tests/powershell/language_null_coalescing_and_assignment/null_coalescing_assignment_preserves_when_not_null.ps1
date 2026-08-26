# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_assignment_preserves_when_not_null
$target = "ExistingValue"
$target ??= "NewValue"
if ($target -ne "ExistingValue") {
    Write-Host "FAIL: ??= when target is non-null should preserve existing value, got '$target'"
    exit 1
}
Write-Host "PASS"
exit 0
