# vybe-test: powershell/collections_arraylist_legacy/toarray_typed_conversion
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(100, 200, 300))
$intArr = $al.ToArray([type]"int")
if ($intArr.GetType().FullName -ne "System.Int32[]" -or $intArr[2] -ne 300) {
    Write-Host "FAIL: Typed ToArray failed"
    exit 1
}
Write-Host "PASS"
exit 0
