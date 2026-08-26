# vybe-test: powershell/json_nested_payload_depth/depth_with_generic_list_nesting
$l2 = [System.Collections.Generic.List[int]]::new([int[]]@(10, 20, 30))
$json = $l2 | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered[2] -ne 30) {
    Write-Host "FAIL: Generic list depth failed"
    exit 1
}
Write-Host "PASS"
exit 0
