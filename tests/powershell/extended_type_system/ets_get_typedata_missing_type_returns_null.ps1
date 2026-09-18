# vybe-test: powershell/extended_type_system/ets_get_typedata_missing_type_returns_null
# Querying Get-TypeData for an unregistered type returns $null without throwing
$td = Get-TypeData -TypeName "Completely.Fabricated.TypeName.ForTesting.Xyz99"

if ($null -ne $td) {
    Write-Host "FAIL: expected `$null for unregistered type, got: $td"
    exit 1
}

Write-Host "PASS"
exit 0
