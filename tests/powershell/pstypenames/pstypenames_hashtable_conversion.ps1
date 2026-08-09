# vybe-test: powershell/pstypenames/pstypenames_hashtable_conversion
$hash = @{ Key = "Value" }
$obj = [pscustomobject]$hash
$obj.psobject.TypeNames.Insert(0, "ConvertedHashtable")
if ($obj.psobject.TypeNames[0] -ne "ConvertedHashtable") {
    Write-Host "FAIL: TypeNames insertion on converted hashtable expected ConvertedHashtable"
    exit 1
}
Write-Host "PASS"
exit 0
