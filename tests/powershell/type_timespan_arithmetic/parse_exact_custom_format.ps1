# vybe-test: powershell/type_timespan_arithmetic/parse_exact_custom_format
$str = "12:34:56"
$ts = [timespan]::ParseExact($str, "hh\:mm\:ss", [System.Globalization.CultureInfo]::InvariantCulture)
if ($ts.Hours -ne 12 -or $ts.Minutes -ne 34 -or $ts.Seconds -ne 56) {
    Write-Host "FAIL: ParseExact failed"
    exit 1
}
Write-Host "PASS"
exit 0
