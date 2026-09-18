# vybe-test: powershell/culture_and_globalization_cmdlets/culture_nonexistent_name_throws_exception
# Querying a non-existent culture name with -ErrorAction Stop throws a CultureNotFoundException
$threwError = $false
try {
    Get-Culture -Name "invalid-locale-9988" -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: non-existent culture name did not throw an error"
    exit 1
}

Write-Host "PASS"
exit 0
