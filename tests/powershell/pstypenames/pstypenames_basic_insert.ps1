# vybe-test: powershell/pstypenames/pstypenames_basic_insert
$obj = [pscustomobject]@{ Id = 1 }
$obj.psobject.TypeNames.Insert(0, "MyCustomType")
if ($obj.psobject.TypeNames[0] -ne "MyCustomType") {
    Write-Host "FAIL: TypeNames.Insert(0, 'MyCustomType') expected 'MyCustomType'"
    exit 1
}
Write-Host "PASS"
exit 0
