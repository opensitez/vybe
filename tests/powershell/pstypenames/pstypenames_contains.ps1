# vybe-test: powershell/pstypenames/pstypenames_contains
$obj = [pscustomobject]@{ Y = 2 }
$obj.psobject.TypeNames.Add("AddedType")
if (-not $obj.psobject.TypeNames.Contains("AddedType")) {
    Write-Host "FAIL: TypeNames.Contains('AddedType') expected true"
    exit 1
}
Write-Host "PASS"
exit 0
