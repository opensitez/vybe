# vybe-test: powershell/psprovider_introspection_engine/psprovider_get_all_returns_providerinfo
# Get-PSProvider returns instances of System.Management.Automation.ProviderInfo
$providers = Get-PSProvider

if ($providers.Count -lt 5) {
    Write-Host "FAIL: expected at least 5 core providers, got $($providers.Count)"
    exit 1
}

if (-not ($providers[0] -is [System.Management.Automation.ProviderInfo])) {
    Write-Host "FAIL: expected ProviderInfo type, got: $($providers[0].GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
