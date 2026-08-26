# vybe-test: powershell/collections_arraylist_legacy/setrange_overwrite
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(1, 2, 3, 4))
$al.SetRange(1, @(20, 30))
if ($al[1] -ne 20 -or $al[2] -ne 30 -or $al.Count -ne 4) {
    Write-Host "FAIL: SetRange failed"
    exit 1
}
Write-Host "PASS"
exit 0
