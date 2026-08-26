# vybe-test: powershell/type_biginteger_arithmetic/conversion_from_byte_array
[byte[]]$bytes = @(2, 1, 0)
$val = [bigint]::new($bytes)
if ($val -ne [bigint]258) {
    Write-Host "FAIL: expected 258 from bytes, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
