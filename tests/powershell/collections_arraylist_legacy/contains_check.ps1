# vybe-test: powershell/collections_arraylist_legacy/contains_check
$al = [System.Collections.ArrayList]::new()
$al.Add(42)
if (-not $al.Contains(42) -or $al.Contains(99)) {
    Write-Host "FAIL: Contains failed"
    exit 1
}
Write-Host "PASS"
exit 0
