# vybe-test: powershell/collections_arraylist_legacy/toarray_conversion
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@("alpha", "beta"))
$arr = $al.ToArray()
if ($arr.Length -ne 2 -or $arr[0] -ne "alpha") {
    Write-Host "FAIL: ToArray failed"
    exit 1
}
Write-Host "PASS"
exit 0
