# vybe-test: powershell/pstypenames/pstypenames_remove
$obj = [pscustomobject]@{ A = 1 }
$obj.psobject.TypeNames.Insert(0, "ToRemoveType")
$obj.psobject.TypeNames.Remove("ToRemoveType")
if ($obj.psobject.TypeNames.Contains("ToRemoveType")) {
    Write-Host "FAIL: TypeNames.Remove('ToRemoveType') failed"
    exit 1
}
Write-Host "PASS"
exit 0
