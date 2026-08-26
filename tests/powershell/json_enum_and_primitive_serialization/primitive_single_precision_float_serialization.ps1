# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_single_precision_float_serialization
[float]$f = 3.14
$json = @{ FloatVal = $f } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ([math]::Abs($recovered.FloatVal - 3.14) -gt 1e-4) {
    Write-Host "FAIL: Float serialization failed, got $($recovered.FloatVal)"
    exit 1
}
Write-Host "PASS"
exit 0
