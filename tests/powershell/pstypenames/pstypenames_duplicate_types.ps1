# vybe-test: powershell/pstypenames/pstypenames_duplicate_types
$obj = [pscustomobject]@{ Val = 1 }
$obj.psobject.TypeNames.Insert(0, "SameType")
$obj.psobject.TypeNames.Insert(0, "SameType")
if ($obj.psobject.TypeNames[0] -ne "SameType" -or $obj.psobject.TypeNames[1] -ne "SameType") {
    Write-Host "FAIL: duplicate TypeNames insertion failed"
    exit 1
}
Write-Host "PASS"
exit 0
