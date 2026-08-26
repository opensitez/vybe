# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_int64_minvalue_serialization
$obj = @{ Min = [int64]::MinValue }
$json = $obj | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Min -ne [int64]::MinValue) {
    Write-Host "FAIL: Int64 MinValue serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
