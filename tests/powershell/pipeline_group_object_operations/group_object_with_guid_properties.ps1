# vybe-test: powershell/pipeline_group_object_operations/group_object_with_guid_properties
$g1 = [guid]::NewGuid()
$g2 = [guid]::NewGuid()
$items = @(
    [pscustomobject]@{ Session = $g1; Val = 1 },
    [pscustomobject]@{ Session = $g2; Val = 2 },
    [pscustomobject]@{ Session = $g1; Val = 3 }
)
$groups = @($items | Group-Object -Property Session)
if ($groups.Count -ne 2) {
    Write-Host "FAIL: Group-Object with GUID properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
