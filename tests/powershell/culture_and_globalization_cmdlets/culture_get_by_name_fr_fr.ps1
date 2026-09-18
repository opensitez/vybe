# vybe-test: powershell/culture_and_globalization_cmdlets/culture_get_by_name_fr_fr
# Get-Culture -Name fr-FR queries the French (France) CultureInfo
$fr = Get-Culture -Name "fr-FR"

if ($fr.Name -ne "fr-FR") {
    Write-Host "FAIL: culture Name mismatch, expected 'fr-FR', got: '$($fr.Name)'"
    exit 1
}

if ($fr.TwoLetterISOLanguageName -ne "fr") {
    Write-Host "FAIL: ISO language mismatch, expected 'fr', got: '$($fr.TwoLetterISOLanguageName)'"
    exit 1
}

Write-Host "PASS"
exit 0
