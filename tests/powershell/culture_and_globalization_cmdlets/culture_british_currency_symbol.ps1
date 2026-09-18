# vybe-test: powershell/culture_and_globalization_cmdlets/culture_british_currency_symbol
# The British (en-GB) culture defines the Pound Sterling '£' as its CurrencySymbol
$gb = Get-Culture -Name "en-GB"
$symbol = $gb.NumberFormat.CurrencySymbol

if ($symbol -ne "£") {
    Write-Host "FAIL: expected British currency symbol '£', got: '$symbol'"
    exit 1
}

Write-Host "PASS"
exit 0
