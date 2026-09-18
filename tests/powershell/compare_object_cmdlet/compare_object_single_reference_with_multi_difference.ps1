# vybe-test: powershell/compare_object_cmdlet/compare_object_single_reference_with_multi_difference
# Comparing a single-element reference array against a multi-element difference array
$diff = Compare-Object @("only_ref") @("diff1", "diff2")

if ($diff.Count -ne 3) {
    Write-Host "FAIL: expected 3 records, got $($diff.Count)"
    exit 1
}

$refItem = $diff | Where-Object { $_.SideIndicator -eq '<=' }
if ($refItem.InputObject -ne "only_ref") {
    Write-Host "FAIL: unexpected reference item: $($refItem.InputObject)"
    exit 1
}

Write-Host "PASS"
exit 0
