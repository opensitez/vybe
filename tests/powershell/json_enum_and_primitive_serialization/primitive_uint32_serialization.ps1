# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_uint32_serialization
[uint32]$u = 3000000000
$json = @{ UIntVal = $u } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.UIntVal -ne 3000000000) {
    Write-Host "FAIL: UInt32 serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
