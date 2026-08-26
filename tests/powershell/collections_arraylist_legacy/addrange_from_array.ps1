# vybe-test: powershell/collections_arraylist_legacy/addrange_from_array
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(10, 20, 30))
if ($al.Count -ne 3 -or $al[2] -ne 30) {
    Write-Host "FAIL: AddRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
