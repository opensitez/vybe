# vybe-test: powershell/null_coalescing_assignment/null_assignment_when_null
$val = $null
$val ??= "Assigned"
if ($val -ne "Assigned") {
    Write-Host "FAIL: ??= when null expected 'Assigned', got '$val'"
    exit 1
}
Write-Host "PASS"
exit 0
