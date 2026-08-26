# vybe-test: powershell/collections_arraylist_legacy/removerange_slice
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(1, 2, 3, 4, 5))
$al.RemoveRange(1, 3)
if ($al.Count -ne 2 -or $al[0] -ne 1 -or $al[1] -ne 5) {
    Write-Host "FAIL: RemoveRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
