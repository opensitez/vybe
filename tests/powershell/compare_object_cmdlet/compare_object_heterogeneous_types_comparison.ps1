# vybe-test: powershell/compare_object_cmdlet/compare_object_heterogeneous_types_comparison
# Compare-Object handles arrays containing mixed primitive types
$diff = Compare-Object @(1, "text") @(1, "other")

if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 diff records across heterogeneous arrays, got $($diff.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
