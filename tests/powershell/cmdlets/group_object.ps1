# vybe-test: powershell/cmdlets/group_object
$items = @(
    [PSCustomObject]@{ Type = "A"; Value = 1 },
    [PSCustomObject]@{ Type = "A"; Value = 2 },
    [PSCustomObject]@{ Type = "B"; Value = 3 }
)
$grouped = $items | Group-Object -Property Type
if ($grouped.Count -ne 2) {
    Write-Host "FAIL: expected 2 groups, got $($grouped.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
