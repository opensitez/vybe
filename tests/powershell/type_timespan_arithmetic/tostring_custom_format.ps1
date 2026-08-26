# vybe-test: powershell/type_timespan_arithmetic/tostring_custom_format
$ts = [timespan]::new(2, 4, 6, 8)
$formatted = $ts.ToString("d\.hh\:mm\:ss")
if ($formatted -ne "2.04:06:08") {
    Write-Host "FAIL: expected '2.04:06:08', got '$formatted'"
    exit 1
}
Write-Host "PASS"
exit 0
