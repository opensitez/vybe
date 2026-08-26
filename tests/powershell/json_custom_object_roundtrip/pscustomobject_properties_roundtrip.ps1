# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_properties_roundtrip
$orig = [pscustomobject]@{
    Name = "Alice"
    Age = 30
    Active = $true
}
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Name -ne "Alice" -or $recovered.Age -ne 30 -or $recovered.Active -ne $true) {
    Write-Host "FAIL: PSCustomObject roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
