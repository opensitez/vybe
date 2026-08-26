# vybe-test: powershell/collections_arraylist_legacy/sort_comparable_elements
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(40, 10, 30, 20))
$al.Sort()
if ($al[0] -ne 10 -or $al[1] -ne 20 -or $al[2] -ne 30 -or $al[3] -ne 40) {
    Write-Host "FAIL: Sort failed"
    exit 1
}
Write-Host "PASS"
exit 0
