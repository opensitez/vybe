# vybe-test: powershell/json_nested_payload_depth/depth_with_generic_dictionary_nesting
$d1 = [System.Collections.Generic.Dictionary[string, string]]::new()
$d1.Add("k1", "v1")
$json = $d1 | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.k1 -ne "v1") {
    Write-Host "FAIL: Generic dictionary serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
