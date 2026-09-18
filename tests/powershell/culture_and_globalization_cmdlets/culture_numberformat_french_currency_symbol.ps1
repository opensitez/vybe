# vybe-test: powershell/culture_and_globalization_cmdlets/culture_numberformat_french_currency_symbol
# The French (fr-FR) culture defines the Euro '€' as its CurrencySymbol
$fr = Get-Culture -Name "fr-FR"
$symbol = $fr.NumberFormat.CurrencySymbol

if ($symbol -ne "€") {
    Write-Host "FAIL: expected Euro currency symbol '€', got: '$symbol'"
    exit 1
}

Write-Host "PASS"
exit 0
