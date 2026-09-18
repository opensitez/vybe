# vybe-test: powershell/culture_and_globalization_cmdlets/culture_japanese_yen_currency_symbol
# The Japanese (ja-JP) culture defines the Yen '¥' as its CurrencySymbol
$jp = Get-Culture -Name "ja-JP"
$symbol = $jp.NumberFormat.CurrencySymbol

if ($symbol -ne "¥") {
    Write-Host "FAIL: expected Japanese Yen symbol '¥', got: '$symbol'"
    exit 1
}

Write-Host "PASS"
exit 0
