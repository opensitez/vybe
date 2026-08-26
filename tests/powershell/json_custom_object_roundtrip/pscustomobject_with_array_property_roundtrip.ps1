# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_array_property_roundtrip
$orig = [pscustomobject]@{ Tags = @("web", "api", "prod") }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Tags.Count -ne 3 -or $recovered.Tags[1] -ne "api") {
    Write-Host "FAIL: Array property roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
