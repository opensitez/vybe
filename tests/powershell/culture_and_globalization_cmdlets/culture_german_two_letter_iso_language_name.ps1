# vybe-test: powershell/culture_and_globalization_cmdlets/culture_german_two_letter_iso_language_name
# The German (de-DE) culture has TwoLetterISOLanguageName of 'de'
$de = Get-Culture -Name "de-DE"

if ($de.TwoLetterISOLanguageName -ne "de") {
    Write-Host "FAIL: expected TwoLetterISOLanguageName 'de', got: '$($de.TwoLetterISOLanguageName)'"
    exit 1
}

Write-Host "PASS"
exit 0
