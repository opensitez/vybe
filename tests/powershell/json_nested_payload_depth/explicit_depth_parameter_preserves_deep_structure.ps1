# vybe-test: powershell/json_nested_payload_depth/explicit_depth_parameter_preserves_deep_structure
$obj = @{
    A = @{
        B = @{
            C = @{
                D = "FoundD"
            }
        }
    }
}
$json = $obj | ConvertTo-Json -Depth 5
$recovered = $json | ConvertFrom-Json
if ($recovered.A.B.C.D -ne "FoundD") {
    Write-Host "FAIL: Explicit depth 5 preservation failed, got '$($recovered.A.B.C.D)'"
    exit 1
}
Write-Host "PASS"
exit 0
