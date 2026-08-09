# vybe-test: powershell/pstypenames/pstypenames_count
$obj = [pscustomobject]@{ X = 1 }
$initialCount = $obj.psobject.TypeNames.Count
$obj.psobject.TypeNames.Insert(0, "ExtraType")
if ($obj.psobject.TypeNames.Count -ne ($initialCount + 1)) {
    Write-Host "FAIL: TypeNames.Count expected $($initialCount + 1), got $($obj.psobject.TypeNames.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
