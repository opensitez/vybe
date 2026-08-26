# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_nested_objects_roundtrip
$orig = [pscustomobject]@{
    User = [pscustomobject]@{ Name = "Bob" }
    Meta = [pscustomobject]@{ Version = 2 }
}
$json = $orig | ConvertTo-Json -Depth 3
$recovered = $json | ConvertFrom-Json
if ($recovered.User.Name -ne "Bob" -or $recovered.Meta.Version -ne 2) {
    Write-Host "FAIL: Nested custom objects roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
