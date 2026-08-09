# vybe-test: powershell/pstypenames/pstypenames_clear
$obj = [pscustomobject]@{ A = 1 }
$obj.psobject.TypeNames.Clear()
if ($obj.psobject.TypeNames.Count -ne 0) {
    Write-Host "FAIL: TypeNames.Clear() expected Count 0, got $($obj.psobject.TypeNames.Count)"
    exit 1
}
Write-Host "PASS"
exit 0
