# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_boolean_true_and_false_serialization
$obj = @{ TrueVal = $true; FalseVal = $false }
$json = $obj | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.TrueVal -ne $true -or $recovered.FalseVal -ne $false) {
    Write-Host "FAIL: Boolean serialization failed"
    exit 1
}
Write-Host "PASS"
exit 0
