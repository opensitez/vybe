# vybe-test: powershell/json_nested_payload_depth/depth_with_boolean_flags_deeply_nested
$tree = @{ Auth = @{ Roles = @{ Admin = $true; Guest = $false } } }
$json = $tree | ConvertTo-Json -Depth 4
$recovered = $json | ConvertFrom-Json
if ($recovered.Auth.Roles.Admin -ne $true -or $recovered.Auth.Roles.Guest -ne $false) {
    Write-Host "FAIL: Deeply nested boolean flags failed"
    exit 1
}
Write-Host "PASS"
exit 0
