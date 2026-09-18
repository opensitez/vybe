# vybe-test: powershell/psprovider_introspection_engine/psprovider_name_array_query
# Get-PSProvider accepts an array of provider names to query multiple providers at once
$providers = @(Get-PSProvider -PSProvider Environment, Variable)

if ($providers.Count -ne 2) {
    Write-Host "FAIL: expected 2 providers from array query, got $($providers.Count)"
    exit 1
}

$names = @($providers | ForEach-Object { $_.Name })
if ($names -notcontains "Environment" -or $names -notcontains "Variable") {
    Write-Host "FAIL: provider names mismatch: @($($names -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
