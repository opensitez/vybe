# vybe-test: powershell/culture_and_globalization_cmdlets/culture_specific_double_parsing_with_numberformat
# Culture-specific NumberFormat enables correct parsing of locale-formatted numeric strings
$fr = Get-Culture -Name "fr-FR"

# French uses comma as decimal separator: "1234,56" = 1234.56
$parsed = [double]::Parse("1234,56", $fr.NumberFormat)

if ($parsed -ne 1234.56) {
    Write-Host "FAIL: French decimal parse mismatch, expected 1234.56, got: $parsed"
    exit 1
}

Write-Host "PASS"
exit 0
