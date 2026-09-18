# vybe-test: powershell/extended_type_system/ets_get_typedata_returns_typedata_instance
# Get-TypeData returns a TypeData metadata object for the specified type
$td = Get-TypeData -TypeName System.DateTime

if ($null -eq $td) {
    Write-Host "FAIL: Get-TypeData returned `$null"
    exit 1
}

if ($td.TypeName -ne "System.DateTime") {
    Write-Host "FAIL: expected TypeName 'System.DateTime', got: '$($td.TypeName)'"
    exit 1
}

Write-Host "PASS"
exit 0
