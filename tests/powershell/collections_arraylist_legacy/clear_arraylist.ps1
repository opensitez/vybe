# vybe-test: powershell/collections_arraylist_legacy/clear_arraylist
$al = [System.Collections.ArrayList]::new()
$al.Add(1); $al.Add(2)
$al.Clear()
if ($al.Count -ne 0) {
    Write-Host "FAIL: Clear failed"
    exit 1
}
Write-Host "PASS"
exit 0
