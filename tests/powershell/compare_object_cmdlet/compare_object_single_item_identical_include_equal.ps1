# vybe-test: powershell/compare_object_cmdlet/compare_object_single_item_identical_include_equal
# Comparing single identical items with -IncludeEqual produces 1 element with '=='
$diff = Compare-Object @("singleton") @("singleton") -IncludeEqual

if ($diff.Count -ne 1) {
    Write-Host "FAIL: expected 1 record, got $($diff.Count)"
    exit 1
}

if ($diff[0].InputObject -ne "singleton" -or $diff[0].SideIndicator -ne '==') {
    Write-Host "FAIL: expected 'singleton' with '==', got: $($diff[0].InputObject) $($diff[0].SideIndicator)"
    exit 1
}

Write-Host "PASS"
exit 0
