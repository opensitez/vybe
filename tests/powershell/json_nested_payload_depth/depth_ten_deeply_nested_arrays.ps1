# vybe-test: powershell/json_nested_payload_depth/depth_ten_deeply_nested_arrays
$current = @{ Leaf = "DeepValue" }
for ($i = 0; $i -lt 5; $i++) {
    $current = @{ Child = $current }
}
$json = $current | ConvertTo-Json -Depth 10
$recovered = $json | ConvertFrom-Json
if ($recovered.Child.Child.Child.Child.Child.Leaf -ne "DeepValue") {
    Write-Host "FAIL: Deeply nested JSON depth failed"
    exit 1
}
Write-Host "PASS"
exit 0
