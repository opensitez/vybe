# vybe-test: powershell/pstypenames/pstypenames_copy
$obj1 = [pscustomobject]@{ Val = 1 }
$obj1.psobject.TypeNames.Insert(0, "CopyType")
$obj2 = [pscustomobject]@{ Val = 2 }
foreach ($t in $obj1.psobject.TypeNames) {
    if (-not $obj2.psobject.TypeNames.Contains($t)) {
        $obj2.psobject.TypeNames.Add($t)
    }
}
if (-not $obj2.psobject.TypeNames.Contains("CopyType")) {
    Write-Host "FAIL: TypeNames copy between objects failed"
    exit 1
}
Write-Host "PASS"
exit 0
