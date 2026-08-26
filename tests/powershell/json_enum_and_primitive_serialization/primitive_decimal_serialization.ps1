# vybe-test: powershell/json_enum_and_primitive_serialization/primitive_decimal_serialization
[decimal]$d = 99.99
$json = @{ DecVal = $d } | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.DecVal -ne 99.99) {
    Write-Host "FAIL: Decimal serialization failed, got $($recovered.DecVal)"
    exit 1
}
Write-Host "PASS"
exit 0
