# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_int16_serialization
[int16]$s = -32000
$json = @{ ShortVal = $s } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.ShortVal -ne -32000) {
    Write-Host "FAIL: Int16 serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
