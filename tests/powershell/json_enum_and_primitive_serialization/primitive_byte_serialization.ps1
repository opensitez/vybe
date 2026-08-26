# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_byte_serialization
[byte]$b = 255
$json = @{ ByteVal = $b } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.ByteVal -ne 255) {
    Write-Host "FAIL: Byte serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
