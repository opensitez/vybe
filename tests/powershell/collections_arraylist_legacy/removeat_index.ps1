# vybe-test: powershell/collections_arraylist_legacy/removeat_index
$al = [System.Collections.ArrayList]::new()
$al.Add(100); $al.Add(200); $al.Add(300)
$al.RemoveAt(1)
if ($al.Count -ne 2 -or $al[1] -ne 300) {
    Write-Host "FAIL: RemoveAt failed"
    exit 1
}
Write-Host "PASS"
exit 0
