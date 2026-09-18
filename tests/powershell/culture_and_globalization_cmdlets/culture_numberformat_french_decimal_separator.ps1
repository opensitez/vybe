# vybe-test: powershell/culture_and_globalization_cmdlets/culture_numberformat_french_decimal_separator
# The French (fr-FR) culture defines a comma ',' as its NumberDecimalSeparator
$fr = Get-Culture -Name "fr-FR"
$sep = $fr.NumberFormat.NumberDecimalSeparator

if ($sep -ne ",") {
    Write-Host "FAIL: expected French decimal separator ',', got: '$sep'"
    exit 1
}

Write-Host "PASS"
exit 0
