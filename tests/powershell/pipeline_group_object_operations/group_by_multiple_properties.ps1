# vybe-test: powershell/pipeline_group_object_operations/group_by_multiple_properties
$items = @(
    [pscustomobject]@{ A = 1; B = "X"; Val = "First" },
    [pscustomobject]@{ A = 1; B = "Y"; Val = "Second" },
    [pscustomobject]@{ A = 1; B = "X"; Val = "Third" }
)
$groups = @($items | Group-Object -Property A, B)
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object multiple properties failed, got $($groups.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
