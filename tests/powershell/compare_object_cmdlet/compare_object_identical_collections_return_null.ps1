# vybe-test: powershell/compare_object_cmdlet/compare_object_identical_collections_return_null
# When reference and difference collections are identical, Compare-Object emits nothing ($null) by default
$diff = Compare-Object @("apple", "banana", "cherry") @("apple", "banana", "cherry")

if ($null -ne $diff) {
    Write-Host "FAIL: expected `$null for identical collections, got: $diff"
    exit 1
}

Write-Host "PASS"
exit 0
