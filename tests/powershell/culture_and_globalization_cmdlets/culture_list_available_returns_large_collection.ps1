# vybe-test: powershell/culture_and_globalization_cmdlets/culture_list_available_returns_large_collection
# Get-Culture -ListAvailable returns a large collection of installed system culture definitions
$cultures = Get-Culture -ListAvailable

if ($cultures.Count -lt 100) {
    Write-Host "FAIL: expected at least 100 available cultures, got: $($cultures.Count)"
    exit 1
}

# All items must be CultureInfo instances
if (-not ($cultures[0] -is [System.Globalization.CultureInfo])) {
    Write-Host "FAIL: expected CultureInfo type in ListAvailable results"
    exit 1
}

Write-Host "PASS"
exit 0
