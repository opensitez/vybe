# vybe-test: powershell/string_culture_case_conversion/number_parse_invariant_culture
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$d = [double]::Parse("9876.54", $ci)
if ($d -ne 9876.54) {
    Write-Host "FAIL: Invariant culture double parse failed"
    exit 1
}
Write-Host "PASS"
exit 0
