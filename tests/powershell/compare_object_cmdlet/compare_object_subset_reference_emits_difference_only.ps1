# vybe-test: powershell/compare_object_cmdlet/compare_object_subset_reference_emits_difference_only
# When Reference is a strict subset of Difference, only the extra Difference items are emitted
$diff = Compare-Object @("item1") @("item1", "item2")

if ($diff.Count -ne 1) {
    Write-Host "FAIL: expected 1 difference element, got $($diff.Count)"
    exit 1
}

if ($diff[0].InputObject -ne "item2" -or $diff[0].SideIndicator -ne '=>') {
    Write-Host "FAIL: expected 'item2' with '=>', got: $($diff[0].InputObject) $($diff[0].SideIndicator)"
    exit 1
}

Write-Host "PASS"
exit 0
