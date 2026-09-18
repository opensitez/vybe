# vybe-test: powershell/culture_and_globalization_cmdlets/culture_get_culture_returns_cultureinfo
# Get-Culture returns an instance of System.Globalization.CultureInfo representing the current session culture
$culture = Get-Culture

if ($null -eq $culture) {
    Write-Host "FAIL: Get-Culture returned `$null"
    exit 1
}

if (-not ($culture -is [System.Globalization.CultureInfo])) {
    Write-Host "FAIL: expected CultureInfo type, got: $($culture.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
