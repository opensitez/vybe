# vybe-test: powershell/string_culture_case_conversion/datetime_parse_exact_invariant
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$dt = [datetime]::ParseExact("2026/08/26", "yyyy/MM/dd", $ci)
if ($dt.Year -ne 2026 -or $dt.Month -ne 8 -or $dt.Day -ne 26) {
    Write-Host "FAIL: ParseExact with InvariantCulture failed"
    exit 1
}
Write-Host "PASS"
exit 0
