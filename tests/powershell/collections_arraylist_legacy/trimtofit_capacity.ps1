# vybe-test: powershell/collections_arraylist_legacy/trimtofit_capacity
$al = [System.Collections.ArrayList]::new(100)
$al.Add(1); $al.Add(2)
$al.TrimToSize()
if ($al.Capacity -ne 2) {
    Write-Host "FAIL: TrimToSize failed"
    exit 1
}
Write-Host "PASS"
exit 0
