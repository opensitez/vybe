# vybe-test: powershell/culture_and_globalization_cmdlets/culture_invariant_culture_empty_name_and_iv_language
# The InvariantCulture has an empty Name string and 'iv' as its TwoLetterISOLanguageName
$inv = [System.Globalization.CultureInfo]::InvariantCulture

if ($inv.Name -ne "") {
    Write-Host "FAIL: expected InvariantCulture Name to be empty string, got: '$($inv.Name)'"
    exit 1
}

if ($inv.TwoLetterISOLanguageName -ne "iv") {
    Write-Host "FAIL: expected TwoLetterISOLanguageName 'iv', got: '$($inv.TwoLetterISOLanguageName)'"
    exit 1
}

Write-Host "PASS"
exit 0
