# vybe-test: powershell/pstypenames/pstypenames_array_input
$list = @([pscustomobject]@{ A = 1 }, [pscustomobject]@{ B = 2 })
$list | ForEach-Object { $_.psobject.TypeNames.Insert(0, "CommonType") }
if ($list[0].psobject.TypeNames[0] -ne "CommonType" -or $list[1].psobject.TypeNames[0] -ne "CommonType") {
    Write-Host "FAIL: array items TypeNames insertion failed"
    exit 1
}
Write-Host "PASS"
exit 0
