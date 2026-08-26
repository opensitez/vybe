# vybe-test: powershell/string_culture_case_conversion/number_format_invariant_decimal_point
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$num = 1234.56
$str = $num.ToString("F2", $ci)
if ($str -ne "1234.56") {
    Write-Host "FAIL: Invariant culture number format expected dot, got $str"
    exit 1
}
Write-Host "PASS"
exit 0
