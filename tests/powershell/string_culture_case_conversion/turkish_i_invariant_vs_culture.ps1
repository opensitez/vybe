# vybe-test: powershell/string_culture_case_conversion/turkish_i_invariant_vs_culture
$ci = [System.Globalization.CultureInfo]::InvariantCulture
$str = "i"
$upper = $str.ToUpper($ci)
if ($upper -ne "I") {
    Write-Host "FAIL: Invariant ToUpper on 'i' expected 'I', got $upper"
    exit 1
}
Write-Host "PASS"
exit 0
