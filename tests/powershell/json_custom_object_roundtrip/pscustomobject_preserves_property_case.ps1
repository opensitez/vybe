# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_preserves_property_case
$orig = [pscustomobject]@{ CamelCase = 1; PascalCase = 2; UPPER = 3 }
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
$names = @($recovered.PSObject.Properties | ForEach-Object { $_.Name })
if (-not ($names -contains "CamelCase") -or -not ($names -contains "PascalCase") -or -not ($names -contains "UPPER")) {
    Write-Host "FAIL: Property casing preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
