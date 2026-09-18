# vybe-test: powershell/extended_type_system/ets_get_typedata_all_types
# Calling Get-TypeData without arguments returns TypeData configurations across all registered types
$allTypeData = Get-TypeData

if ($null -eq $allTypeData) {
    Write-Host "FAIL: Get-TypeData returned `$null"
    exit 1
}

if ($allTypeData.Count -lt 10) {
    Write-Host "FAIL: expected broad collection of registered TypeData configurations, got count $($allTypeData.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
