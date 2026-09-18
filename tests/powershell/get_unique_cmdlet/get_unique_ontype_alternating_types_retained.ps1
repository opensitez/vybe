# vybe-test: powershell/get_unique_cmdlet/get_unique_ontype_alternating_types_retained
# With -OnType, non-consecutive occurrences of the same type are retained
$items = @(1, "alpha", 2, "beta")
$unique = @($items | Get-Unique -OnType)

if ($unique.Count -ne 4) {
    Write-Host "FAIL: alternating types should all be retained, got $($unique.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
