# vybe-test: powershell/type_timespan_arithmetic/parse_standard_d_hh_mm_ss
$str = "3.12:30:45"
$ts = [timespan]::Parse($str)
if ($ts.Days -ne 3 -or $ts.Hours -ne 12 -or $ts.Minutes -ne 30 -or $ts.Seconds -ne 45) {
    Write-Host "FAIL: parsed mismatch for $str"
    exit 1
}
Write-Host "PASS"
exit 0
