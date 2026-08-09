# vybe-test: powershell/type_converters/type_converter_invalid_throw
try {
    [int]$invalid = "NotANumber"
    Write-Host "FAIL: invalid string to int conversion expected throw"
    exit 1
} catch {
    Write-Host "PASS"
    exit 0
}
