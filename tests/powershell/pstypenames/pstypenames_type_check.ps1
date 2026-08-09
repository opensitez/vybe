# vybe-test: powershell/pstypenames/pstypenames_type_check
$obj = [pscustomobject]@{ Status = "OK" }
$obj.psobject.TypeNames.Insert(0, "Vybe.System.StatusObject")
if (-not ($obj.psobject.TypeNames -contains "Vybe.System.StatusObject")) {
    Write-Host "FAIL: TypeNames contains check failed for inserted type"
    exit 1
}
Write-Host "PASS"
exit 0
