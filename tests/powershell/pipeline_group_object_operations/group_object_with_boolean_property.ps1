# vybe-test: powershell/pipeline_group_object_operations/group_object_with_boolean_property
$items = @(
    [pscustomobject]@{ Active = $true; Name = "A" },
    [pscustomobject]@{ Active = $false; Name = "B" },
    [pscustomobject]@{ Active = $true; Name = "C" }
)
$groups = @($items | Group-Object -Property Active)
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object with boolean property failed"
    exit 1
}
Write-Host "PASS"
exit 0
