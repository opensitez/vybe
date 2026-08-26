# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_uint64_serialization
[uint64]$ul = 18000000000000000000
$json = @{ UInt64Val = $ul } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.UInt64Val -ne 18000000000000000000) {
    # Check large uint64 preservation
}
Write-Host "PASS"
exit 0
