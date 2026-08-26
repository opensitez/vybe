# vybe-test: powershell/math_floating_point_epsilon/double_tostring_roundtrip_format_r
$orig = 0.12345678901234567
$str = $orig.ToString("R", [System.Globalization.CultureInfo]::InvariantCulture)
$parsed = [double]::Parse($str, [System.Globalization.CultureInfo]::InvariantCulture)
if ($orig -ne $parsed) {
    Write-Host "FAIL: 'R' roundtrip format failed"
    exit 1
}
Write-Host "PASS"
exit 0
