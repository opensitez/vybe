# vybe-test: powershell/psprovider_introspection_engine/psprovider_nonexistent_provider_throws
# Querying a non-existent provider name throws an error with -ErrorAction Stop
$threwError = $false
try {
    Get-PSProvider "CompletelyFabricatedProvider_9988" -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: non-existent provider query did not throw"
    exit 1
}

Write-Host "PASS"
exit 0
