# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_assignment_mutates_when_null
$target = $null
$target ??= "AssignedDefault"
if ($target -ne "AssignedDefault") {
    Write-Host "FAIL: ??= when target is null failed, got '$target'"
    exit 1
}
Write-Host "PASS"
exit 0
