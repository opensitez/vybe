# vybe-test: powershell/collections_arraylist_legacy/foreach_iteration
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(1, 2, 3))
$sum = 0
foreach ($item in $al) { $sum += $item }
if ($sum -ne 6) {
    Write-Host "FAIL: Foreach on ArrayList failed"
    exit 1
}
Write-Host "PASS"
exit 0
