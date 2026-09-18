# vybe-test: powershell/new_object_cmdlet/new_object_psobject_get_type_returns_pscustomobject
# New-Object PSObject produces an object whose type name is PSCustomObject
$obj = New-Object PSObject -Property @{ X = 1 }
$typeName = $obj.GetType().Name

if ($typeName -ne "PSCustomObject") {
    Write-Host "FAIL: expected PSCustomObject, got '$typeName'"
    exit 1
}

Write-Host "PASS"
exit 0
