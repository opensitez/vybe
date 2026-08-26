# vybe-test: powershell/json_nested_payload_depth/deeply_nested_mixed_arrays_and_objects
$tree = @{
    Nodes = @(
        @{ Id = 1; Children = @( @{ Id = 11 }, @{ Id = 12 } ) },
        @{ Id = 2; Children = @( @{ Id = 21 } ) }
    )
}
$json = $tree | ConvertTo-Json -Depth 6
$recovered = $json | ConvertFrom-Json
if ($recovered.Nodes[0].Children[1].Id -ne 12 -or $recovered.Nodes[1].Children[0].Id -ne 21) {
    Write-Host "FAIL: Deeply nested mixed tree failed"
    exit 1
}
Write-Host "PASS"
exit 0
