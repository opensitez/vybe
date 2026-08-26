# vybe-test: powershell/json_custom_object_roundtrip/pscustomobject_with_numeric_data_types
$orig = [pscustomobject]@{
    IntVal = 42
    DoubleVal = 3.14159
    LongVal = 9000000000000
}
$json = $orig | ConvertTo-Json
$recovered = $json | ConvertFrom-Json
if ($recovered.IntVal -ne 42 -or [math]::Abs($recovered.DoubleVal - 3.14159) -gt 1e-5 -or $recovered.LongVal -ne 9000000000000) {
    Write-Host "FAIL: Numeric data types roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
