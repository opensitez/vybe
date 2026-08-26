# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_zero_and_negative_zero_serialization
$obj = @{ Zero = 0.0; NegZero = -0.0 }
$json = $obj | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.Zero -ne 0.0 -or $recovered.NegZero -ne 0.0) {
    Write-Host "FAIL: Zero and negative zero serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
