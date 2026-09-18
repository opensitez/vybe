# vybe-test: powershell/intrinsic_collection_overloads/foreach_property_extraction_on_missing_property_returns_null
# Heterogeneous object collection where some objects have the property and others do not
$objects = @(
    [pscustomobject]@{ Tag = "A"; Priority = 1 },
    [pscustomobject]@{ Tag = "B" },  # Priority is missing
    [pscustomobject]@{ Tag = "C"; Priority = 3 }
)

# .ForEach('PropertyName') extracts $null for objects missing the property without throwing
$priorities = $objects.ForEach('Priority')

if ($null -eq $priorities -or $priorities.Count -ne 3) {
    Write-Host "FAIL: .ForEach('Priority') expected collection of 3 items, got $($priorities.Count)"
    exit 1
}

if ($priorities[0] -ne 1 -or $null -ne $priorities[1] -or $priorities[2] -ne 3) {
    Write-Host "FAIL: expected @(1, `$null, 3), got @($($priorities -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
