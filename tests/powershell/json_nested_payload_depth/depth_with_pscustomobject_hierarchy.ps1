# vybe-test: powershell/json_nested_payload_depth/depth_with_pscustomobject_hierarchy
$root = [pscustomobject]@{
    Level1 = [pscustomobject]@{
        Level2 = [pscustomobject]@{
            Value = "Deep"
        }
    }
}
$json = $root | ConvertTo-Json -Depth 4
$recovered = $json | ConvertFrom-Json
if ($recovered.Level1.Level2.Value -ne "Deep") {
    Write-Host "FAIL: PSCustomObject hierarchy depth serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
