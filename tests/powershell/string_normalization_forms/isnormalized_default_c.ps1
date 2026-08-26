# vybe-test: powershell/string_normalization_forms/isnormalized_default_c
$str = "Hello"
if (-not $str.IsNormalized()) {
    Write-Host "FAIL: ASCII string should be normalized form C"
    exit 1
}
Write-Host "PASS"
exit 0
