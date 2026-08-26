# vybe-test: powershell/json_nested_payload_depth/depth_pipeline_roundtrip_equality
$orig = @{ A = @{ B = @{ C = 100 } } }
$res = $orig | ConvertTo-Json -Depth 5 | ConvertFrom-Json -AsHashtable
if ($res["A"]["B"]["C"] -ne 100) {
    Write-Host "FAIL: Deep JSON hashtable roundtrip equality failed"
    exit 1
}
Write-Host "PASS"
exit 0
