# vybe-test: powershell/compare_object_cmdlet/compare_object_duplicate_elements_count_discrepancy
# Compare-Object detects discrepancies in duplicate element frequencies between collections
$diff = Compare-Object @("dup", "dup") @("dup")

if ($diff.Count -ne 1) {
    Write-Host "FAIL: expected 1 discrepancy record for duplicate frequency mismatch, got $($diff.Count)"
    exit 1
}

if ($diff.SideIndicator -ne '<=') {
    Write-Host "FAIL: expected '<=' for excess item in reference collection, got: $($diff.SideIndicator)"
    exit 1
}

Write-Host "PASS"
exit 0
