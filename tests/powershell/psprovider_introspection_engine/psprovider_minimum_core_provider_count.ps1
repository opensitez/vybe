# vybe-test: powershell/psprovider_introspection_engine/psprovider_minimum_core_provider_count
# PowerShell ships with at least five core providers: Alias, Environment, FileSystem, Function, and Variable
$providers = @(Get-PSProvider)

if ($providers.Count -lt 5) {
    Write-Host "FAIL: fewer than 5 providers found, got: $($providers.Count)"
    exit 1
}

$names = @($providers | ForEach-Object { $_.Name })
$expectedCore = @("Alias", "Environment", "FileSystem", "Function", "Variable")

foreach ($core in $expectedCore) {
    if ($names -notcontains $core) {
        Write-Host "FAIL: core provider '$core' missing from registered providers: @($($names -join ', '))"
        exit 1
    }
}

Write-Host "PASS"
exit 0
