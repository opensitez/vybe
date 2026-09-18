# vybe-test: powershell/culture_and_globalization_cmdlets/culture_name_lookup_is_case_insensitive
# Get-Culture -Name performs case-insensitive BCP-47 culture name matching
$lowerCase = Get-Culture -Name "de-de"

if ($lowerCase.Name -ne "de-DE") {
    Write-Host "FAIL: lowercase culture name lookup did not normalize to 'de-DE', got: '$($lowerCase.Name)'"
    exit 1
}

Write-Host "PASS"
exit 0
