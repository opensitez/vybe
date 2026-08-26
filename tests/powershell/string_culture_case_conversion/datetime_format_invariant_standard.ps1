# vybe-test: powershell/string_culture_case_conversion/datetime_format_invariant_standard
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$dt = [datetime]::new(2026, 8, 26, 15, 30, 0)
$str = $dt.ToString("yyyy-MM-dd HH:mm:ss", $ci)
if ($str -ne "2026-08-26 15:30:00") {
    Write-Host "FAIL: Invariant datetime format failed"
    exit 1
}
Write-Host "PASS"
exit 0
