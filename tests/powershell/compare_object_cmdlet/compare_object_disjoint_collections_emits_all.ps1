# vybe-test: powershell/compare_object_cmdlet/compare_object_disjoint_collections_emits_all
# When sets share no elements, all items are emitted with appropriate side indicators
$diff = Compare-Object @("a", "b") @("c", "d")

if ($diff.Count -ne 4) {
    Write-Host "FAIL: expected 4 differences for disjoint sets, got $($diff.Count)"
    exit 1
}

$leftCount  = @($diff | Where-Object { $_.SideIndicator -eq '<=' }).Count
$rightCount = @($diff | Where-Object { $_.SideIndicator -eq '=>' }).Count

if ($leftCount -ne 2 -or $rightCount -ne 2) {
    Write-Host "FAIL: expected 2 '<=' and 2 '=>', got left=$leftCount, right=$rightCount"
    exit 1
}

Write-Host "PASS"
exit 0
