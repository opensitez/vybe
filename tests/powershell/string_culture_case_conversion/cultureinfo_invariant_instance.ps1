# vybe-test: powershell/string_culture_case_conversion/cultureinfo_invariant_instance
$ci = [System.Globalization.CultureInfo]::InvariantCulture
if ($ci.Name -ne "" -or $ci.EnglishName -ne "Invariant Language (Invariant Country)") {
    Write-Host "FAIL: InvariantCulture instance check failed"
    exit 1
}
Write-Host "PASS"
exit 0
