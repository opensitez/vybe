# vybe-test: powershell/string_culture_case_conversion/culture_specific_currencysymbol_check
$ci = [System.Globalization.CultureInfo]::InvariantCulture
if ($ci.NumberFormat.CurrencySymbol -ne "¤") {
    Write-Host "FAIL: Invariant currency symbol expected ¤, got $($ci.NumberFormat.CurrencySymbol)"
    exit 1
}
Write-Host "PASS"
exit 0
