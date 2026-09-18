# vybe-test: powershell/culture_and_globalization_cmdlets/culture_multiple_names_array_query
# Get-Culture -Name accepts an array of culture name strings and returns multiple CultureInfo objects
$cultures = @(Get-Culture -Name en-US, fr-FR)

if ($cultures.Count -ne 2) {
    Write-Host "FAIL: expected 2 cultures from array query, got $($cultures.Count)"
    exit 1
}

$names = @($cultures | ForEach-Object { $_.Name })
if ($names -notcontains "en-US" -or $names -notcontains "fr-FR") {
    Write-Host "FAIL: expected 'en-US' and 'fr-FR' in results: @($($names -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
