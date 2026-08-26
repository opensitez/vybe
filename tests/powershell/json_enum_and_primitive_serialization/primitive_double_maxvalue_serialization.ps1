# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_double_maxvalue_serialization
$obj = @{ Max = [double]::MaxValue }
$json = $obj | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Max -le 1e308) {
    Write-Host "FAIL: Double MaxValue serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
