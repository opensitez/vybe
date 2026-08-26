# vybe-test: powershell/collections_arraylist_legacy/getrange_sublist
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(0, 1, 2, 3, 4, 5))
$sub = $al.GetRange(2, 3) # 2, 3, 4
if ($sub.Count -ne 3 -or $sub[0] -ne 2 -or $sub[2] -ne 4) {
    Write-Host "FAIL: GetRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
