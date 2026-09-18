# vybe-test: powershell/compare_object_cmdlet/compare_object_order_independence_by_default
# By default, Compare-Object treats collections with the same items in different order as equal
$diff = Compare-Object @(1, 2, 3) @(3, 1, 2)

if ($null -ne $diff) {
    Write-Host "FAIL: differently ordered collections unexpectedly flagged as differing: $diff"
    exit 1
}

Write-Host "PASS"
exit 0
