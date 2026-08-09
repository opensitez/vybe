# vybe-test: powershell/null_coalescing_assignment/null_assignment_when_not_null
$val = "Existing"
$val ??= "Fallback"
if ($val -ne "Existing") {
    Write-Host "FAIL: ??= when not null expected 'Existing', got '$val'"
    exit 1
}
Write-Host "PASS"
exit 0
